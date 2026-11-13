/**
 * API do sistema de convite de oradores — Discurso Público.
 *
 * Fica "colado" na Google Planilha (Extensões → Apps Script), junto com o arquivo
 * Admin.html, e tem DUAS implantações de App da Web (mesmo código):
 *
 *   1) PÚBLICA (página do orador) — Executar como: Eu | Acesso: Qualquer pessoa
 *        GET  → datas DISPONÍVEIS (de hoje em diante), lista de temas e textos da aba Config.
 *        POST → grava a inscrição na aba Inscricoes e marca a data como RESERVADA.
 *
 *   2) ADMINISTRAÇÃO — Executar como: Usuário que acessa o app
 *                      Acesso: Qualquer pessoa com Conta do Google
 *        GET ?p=admin → tela de administração (Admin.html).
 *        Só entra quem (a) tem acesso de edição à planilha E (b) está na lista
 *        ADMINISTRADORES da aba Config.
 *
 * A reserva usa LockService: se dois oradores enviarem a mesma data ao mesmo tempo,
 * só o primeiro fica com ela; o segundo recebe DATA_INDISPONIVEL e escolhe outra.
 */

const ABA_DATAS = 'Datas';
const ABA_TEMAS = 'Temas';
const ABA_INSCRICOES = 'Inscricoes';
const ABA_CONFIG = 'Config';

// Colunas da aba Datas (base 0)
const D_DATA = 0, D_HORARIO = 1, D_SITUACAO = 2, D_OBS = 3;

// Colunas da aba Inscricoes (base 0)
const I = {
  ID: 0, RECEBIDO: 1, DATA: 2, HORARIO: 3, NOME: 4, CONGREGACAO: 5, WHATSAPP: 6, EMAIL: 7,
  TEMAS_NUM: 8, TEMAS_TXT: 9, TEMA_ESCOLHIDO: 10, HOSPITALIDADE: 11, STATUS: 12, CONVITE: 13,
  CONFIRMADO_POR: 14, OBS: 15, GMAIL: 16, CONFIRMADO_EM: 17, TIPO: 18,
};
// Colunas acrescentadas depois da primeira versão: o script cria o cabeçalho se faltar.
const CABECALHOS_NOVOS = { GMAIL: 'E-MAIL PESSOAL', CONFIRMADO_EM: 'CONFIRMADO EM', TIPO: 'TIPO' };

// Tipos de linha na aba Inscricoes (coluna TIPO; vazio = discurso público).
const TIPO_DISCURSO = 'Discurso público';
const TIPOS_EVENTO = [TIPO_DISCURSO, 'Discurso especial', 'Assembleia de circuito', 'Congresso regional',
  'Reunião especial (Corpo Governante)', 'Visita do superintendente de circuito', 'Outro evento'];

const SIT_DISPONIVEL = 'DISPONIVEL'; // comparado sem acento/maiúsculas (ver norm_)
const SIT_RESERVADA = 'RESERVADA';
const ST_PENDENTE = 'Pendente', ST_CONFIRMADO = 'Confirmado', ST_CANCELADO = 'Cancelado',
  ST_SUBSTITUIDO = 'Substituído';

// Chaves da aba Config que podem ir para a página pública.
const CONFIG_PUBLICA = {
  'TITULO': 'titulo',
  'CONGREGACAO ANFITRIA': 'congregacao',
  'ENDERECO': 'endereco',
  'LOCALIZACAO': 'mapa',
};
const CFG_HORARIO_PADRAO = 'HORARIO PADRAO';
const CFG_URL_PAGINA = 'URL DA PAGINA';
const CFG_ADMINS = 'ADMINISTRADORES';
const CFG_EMAIL_AVISO = 'E-MAIL PARA AVISOS';

// ============================================================ página pública

function doGet(e) {
  if (e && e.parameter && e.parameter.p === 'admin') return paginaAdmin_();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const config = lerConfig_(ss);
    const hoje = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

    const datas = linhas_(ss, ABA_DATAS)
      .filter(r => norm_(r[D_SITUACAO]) === SIT_DISPONIVEL)
      .map(r => ({ data: iso_(r[D_DATA], tz), horario: horario_(r[D_HORARIO], tz, config) }))
      .filter(d => d.data && d.data >= hoje)
      .sort((a, b) => a.data < b.data ? -1 : 1);

    const publica = {};
    Object.keys(CONFIG_PUBLICA).forEach(k => { publica[CONFIG_PUBLICA[k]] = config[k] || ''; });

    return json_({ ok: true, datas: datas, temas: lerTemas_(ss), config: publica });
  } catch (err) {
    return json_({ ok: false, erro: 'ERRO_INTERNO', mensagem: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d.site) return json_({ ok: true }); // campo-isca anti-robô preenchido: ignora em silêncio

    const nome = texto_(d.nome, 120);
    const congregacao = texto_(d.congregacao, 120);
    const whatsapp = texto_(d.whatsapp, 30);
    const email = texto_(d.email, 120);
    const gmail = emailPessoal_(d.gmail);
    const data = String(d.data || '');
    const numeros = Array.isArray(d.temas)
      ? [...new Set(d.temas.map(Number).filter(n => Number.isInteger(n) && n > 0))]
      : [];

    if (!nome || !congregacao || !/^\d{4}-\d{2}-\d{2}$/.test(data) || !numeros.length) {
      return json_({ ok: false, erro: 'DADOS_INVALIDOS',
        mensagem: 'Preencha nome, congregação, a data e pelo menos um tema.' });
    }

    lock.waitLock(20000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const config = lerConfig_(ss);
    const abaDatas = ss.getSheetByName(ABA_DATAS);
    const valores = abaDatas.getDataRange().getValues();
    const i = valores.findIndex((r, idx) => idx > 0 && iso_(r[D_DATA], tz) === data);
    if (i < 0 || norm_(valores[i][D_SITUACAO]) !== SIT_DISPONIVEL) {
      return json_({ ok: false, erro: 'DATA_INDISPONIVEL',
        mensagem: 'Essa data acabou de ser escolhida por outro orador. Escolha outra, por favor.' });
    }

    const titulos = mapaTemas_(ss);
    const temas = numeros.filter(n => titulos[n]).sort((a, b) => a - b);
    if (!temas.length) {
      return json_({ ok: false, erro: 'DADOS_INVALIDOS', mensagem: 'Nenhum tema válido foi marcado.' });
    }

    abaDatas.getRange(i + 1, D_SITUACAO + 1, 1, 2).setValues([[SIT_RESERVADA, nome + ' — ' + congregacao]]);

    const id = Utilities.getUuid().slice(0, 8).toUpperCase();
    const horario = horario_(valores[i][D_HORARIO], tz, config);
    const abaInsc = ss.getSheetByName(ABA_INSCRICOES);
    garantirCabecalhos_(abaInsc);
    const linha = new Array(Object.keys(I).length).fill('');
    linha[I.ID] = id;
    linha[I.RECEBIDO] = new Date();
    linha[I.DATA] = valores[i][D_DATA];
    linha[I.HORARIO] = horario ? "'" + horario : '';
    linha[I.NOME] = nome;
    linha[I.CONGREGACAO] = congregacao;
    linha[I.WHATSAPP] = whatsapp;
    linha[I.EMAIL] = email;
    linha[I.GMAIL] = gmail;
    linha[I.TEMAS_NUM] = "'" + temas.join(', ');
    linha[I.TEMAS_TXT] = temas.map(n => n + ' - ' + titulos[n]).join('\n');
    linha[I.STATUS] = ST_PENDENTE;
    abaInsc.appendRow(linha);
    const n = abaInsc.getLastRow();
    abaInsc.getRange(n, I.RECEBIDO + 1).setNumberFormat('dd/MM/yyyy HH:mm');
    abaInsc.getRange(n, I.DATA + 1).setNumberFormat('dd/MM/yyyy');
    SpreadsheetApp.flush();
    lock.releaseLock();

    avisar_(ss, config, { data, horario, nome, congregacao, whatsapp, email, gmail, temas, titulos });
    return json_({ ok: true, id: id, data: data });
  } catch (err) {
    return json_({ ok: false, erro: 'ERRO_INTERNO', mensagem: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) { /* já liberado */ }
  }
}

/** E-mail opcional para os responsáveis a cada nova inscrição (aba Config). */
function avisar_(ss, config, r) {
  const destino = config[CFG_EMAIL_AVISO];
  if (!destino) return;
  try {
    MailApp.sendEmail({
      to: destino,
      subject: 'Nova inscrição de orador — ' + dataCurta_(r.data) + ' — ' + r.nome,
      body: [
        'Data: ' + dataPorExtenso_(r.data) + (r.horario ? ', às ' + r.horario : ''),
        'Orador: ' + r.nome,
        'Congregação: ' + r.congregacao,
        'WhatsApp: ' + (r.whatsapp || '—'),
        'E-mail: ' + (r.email || '—'),
        'E-mail pessoal: ' + (r.gmail || '—'),
        '',
        'Temas preparados:',
        r.temas.map(n => '  ' + n + ' - ' + r.titulos[n]).join('\n'),
        '',
        'Confira e confirme na tela de administração.',
        'Planilha: ' + ss.getUrl(),
      ].join('\n'),
    });
  } catch (err) {
    console.warn('Falha ao enviar aviso por e-mail: ' + err);
  }
}

// ============================================================ administração
// Funções chamadas pela tela Admin.html via google.script.run.
// Todas começam com exigirAdmin_() — nada é lido nem gravado sem permissão.

function paginaAdmin_() {
  try {
    exigirAdmin_();
  } catch (err) {
    return HtmlService.createHtmlOutput(
      '<div style="font:16px system-ui;padding:32px;max-width:480px;margin:auto">' +
      '<h2>Acesso negado</h2><p>' + escHtml_(err.message) + '</p>' +
      '<p>Peça ao responsável para incluir seu e-mail na lista ADMINISTRADORES da planilha ' +
      'e compartilhar a planilha com você como Editor.</p></div>'
    ).setTitle('Acesso negado');
  }
  return HtmlService.createHtmlOutputFromFile('Admin')
    .setTitle('Administração — Discurso Público')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Lista as inscrições (mais próximas primeiro) com o que a tela precisa. */
function adminListar() {
  const usuario = exigirAdmin_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const titulos = mapaTemas_(ss);
  const lista = linhas_(ss, ABA_INSCRICOES)
    .filter(r => r[I.ID])
    .map(r => {
      const temas = String(r[I.TEMAS_NUM]).split(/[^\d]+/).filter(Boolean).map(Number);
      return {
        id: String(r[I.ID]),
        recebido: r[I.RECEBIDO] instanceof Date ? Utilities.formatDate(r[I.RECEBIDO], tz, 'dd/MM/yyyy HH:mm') : '',
        data: iso_(r[I.DATA], tz),
        horario: String(r[I.HORARIO] || ''),
        nome: String(r[I.NOME]),
        congregacao: String(r[I.CONGREGACAO]),
        whatsapp: String(r[I.WHATSAPP]),
        email: String(r[I.EMAIL]),
        gmail: String(r[I.GMAIL] || ''),
        temas: temas.map(n => ({ numero: n, titulo: titulos[n] || '(tema não encontrado)' })),
        temaEscolhido: numeroTema_(r[I.TEMA_ESCOLHIDO]),
        temaTexto: String(r[I.TEMA_ESCOLHIDO] || '').replace(/^\s*\d+\s*-\s*/, ''),
        hospitalidade: String(r[I.HOSPITALIDADE] || ''),
        status: String(r[I.STATUS] || ST_PENDENTE),
        convite: String(r[I.CONVITE] || ''),
        confirmadoPor: String(r[I.CONFIRMADO_POR] || ''),
        obs: String(r[I.OBS] || ''),
        confirmadoEm: r[I.CONFIRMADO_EM] instanceof Date
          ? Utilities.formatDate(r[I.CONFIRMADO_EM], tz, 'dd/MM/yyyy HH:mm') : '',
        tipo: String(r[I.TIPO] || '') || TIPO_DISCURSO,
      };
    })
    .sort((a, b) => a.data < b.data ? -1 : a.data > b.data ? 1 : 0);
  return { usuario: usuario, inscricoes: lista, ultimas: ultimasApresentacoes_(lista), catalogo: lerTemas_(ss),
    tipos: TIPOS_EVENTO };
}

/** Confirma a inscrição com o tema escolhido e devolve a mensagem/link do WhatsApp. */
function adminConfirmar(id, numeroTema, hospitalidade) {
  const usuario = exigirAdmin_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const { aba, linha, r } = acharInscricao_(ss, id);
  const marcados = String(r[I.TEMAS_NUM]).split(/[^\d]+/).filter(Boolean).map(Number);
  numeroTema = Number(numeroTema);
  if (!marcados.includes(numeroTema)) throw new Error('Escolha um dos temas que o orador marcou.');
  if (String(r[I.STATUS]) !== ST_PENDENTE) throw new Error('Esta inscrição não está mais pendente.');

  const titulo = mapaTemas_(ss)[numeroTema];
  aba.getRange(linha, I.TEMA_ESCOLHIDO + 1).setValue(numeroTema + ' - ' + titulo);
  aba.getRange(linha, I.HOSPITALIDADE + 1).setValue(texto_(hospitalidade, 300));
  aba.getRange(linha, I.STATUS + 1).setValue(ST_CONFIRMADO);
  aba.getRange(linha, I.CONFIRMADO_POR + 1).setValue(usuario);
  registrarConfirmacao_(aba, linha);
  SpreadsheetApp.flush();
  const resposta = adminMensagem(id);
  resposta.email = enviarEmailConvite_(ss, id);
  return resposta;
}

/** Grava a data/hora da confirmação (coluna CONFIRMADO EM). */
function registrarConfirmacao_(aba, linha) {
  garantirCabecalhos_(aba);
  aba.getRange(linha, I.CONFIRMADO_EM + 1).setValue(new Date()).setNumberFormat('dd/MM/yyyy HH:mm');
}

/** Reenvia o convite por e-mail (botão na Gestão). */
function adminEnviarEmail(id) {
  exigirAdmin_();
  return enviarEmailConvite_(SpreadsheetApp.getActiveSpreadsheet(), id);
}

/** Monta (de novo) a mensagem do WhatsApp e o link do convite de uma inscrição confirmada. */
function adminMensagem(id) {
  exigirAdmin_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const { r } = acharInscricao_(ss, id);
  const { nome, data, horario, numero, titulo, anfitria, endereco, mapa, linkConvite } = dadosConvite_(ss, r);

  // Sem emojis: no WhatsApp Desktop/Web eles chegavam como "?" pelo link. Rótulos em negrito
  // (*...*) aparecem certos em qualquer aparelho.
  const texto = [
    'Olá, ' + primeiroNome_(nome) + '! Tudo bem?',
    '',
    'Seu discurso público está *confirmado*. Muito obrigado pela disposição!',
    '',
    '*Data:* ' + capitalizar_(dataPorExtenso_(data)) + (horario ? ', às ' + horario.replace(':', 'h') : ''),
    '*Tema nº ' + numero + ':* _' + titulo + '_',
    anfitria ? '*Local:* Congregação ' + anfitria.replace(/^congrega[cç][aã]o\s+/i, '') : '',
    endereco ? endereco : '',
    mapa ? '*Como chegar:* ' + mapa : '',
    r[I.HOSPITALIDADE] ? '*Hospitalidade:* ' + r[I.HOSPITALIDADE] : '',
    linkConvite ? '\n*Seu convite* (com botões para baixar em PDF e adicionar à agenda):\n' + linkConvite : '',
  ].filter(l => l !== '').join('\n');

  // api.whatsapp.com/send em vez de wa.me: o redirecionamento do wa.me estragava a acentuação/símbolos.
  const fone = foneWhatsApp_(r[I.WHATSAPP]);
  return {
    texto: texto,
    linkConvite: linkConvite,
    linkWhatsApp: fone ? 'https://api.whatsapp.com/send?phone=' + fone + '&text=' + encodeURIComponent(texto) : '',
  };
}

/** Tudo o que o convite precisa, a partir de uma linha da aba Inscricoes já com tema escolhido. */
function dadosConvite_(ss, r) {
  const tz = ss.getSpreadsheetTimeZone();
  const config = lerConfig_(ss);
  const numero = numeroTema_(r[I.TEMA_ESCOLHIDO]);
  if (!numero) throw new Error('Ainda não há tema escolhido para esta inscrição.');
  const c = {
    numero: numero,
    titulo: mapaTemas_(ss)[numero] || '',
    data: iso_(r[I.DATA], tz),
    horario: horario_(r[I.HORARIO], tz, config),
    nome: String(r[I.NOME]).replace(/^'/, ''),
    anfitria: config['CONGREGACAO ANFITRIA'] || '',
    endereco: config['ENDERECO'] || '',
    mapa: /^https:\/\//.test(config['LOCALIZACAO'] || '') ? config['LOCALIZACAO'] : '',
    congregacaoOrador: String(r[I.CONGREGACAO] || '').replace(/^'/, ''),
    linkConvite: '',
  };
  // Dados do convite vão no "#" do link: não passam por servidor nenhum (nem do GitHub).
  // Contatos e telefone ficam na aba Config, para não ficarem públicos no código do GitHub.
  const base = config[CFG_URL_PAGINA];
  if (base) {
    const dados = { n: c.nome, oc: c.congregacaoOrador, d: c.data, h: c.horario, t: c.numero, tt: c.titulo,
      c: c.anfitria, e: c.endereco, m: c.mapa, ce: config['E-MAIL DA CONGREGACAO'] || '',
      ct: String(config['CONTATOS'] || '').split(/\s*\n\s*|\s*\|\s*/).filter(Boolean),
      tel: String(config['TELEFONE PARA CONTATO'] || '').replace(/\D/g, '') };
    // Sem os "=" do final: alguns apps não incluem esses sinais no link clicável.
    c.linkConvite = base.replace(/\/?(index\.html)?$/, '/') + 'convite.html#' +
      Utilities.base64EncodeWebSafe(JSON.stringify(dados), Utilities.Charset.UTF_8).replace(/=+$/, '');
  }
  return c;
}

/**
 * Envia o convite para o e-mail pessoal do orador (Gmail, Hotmail, Outlook…), se informado, com
 * botões "Adicionar ao Google Agenda" e "Adicionar ao Outlook/Hotmail" e um arquivo .ics anexo
 * (serve para qualquer agenda, inclusive a do iPhone).
 * O e-mail sai da conta Google de quem confirmou. Devolve o endereço usado, ou '' se não enviou.
 */
function enviarEmailConvite_(ss, id) {
  const { r } = acharInscricao_(ss, id);
  const para = emailPessoal_(r[I.GMAIL]);
  if (!para) return '';
  const c = dadosConvite_(ss, r);
  const tz = ss.getSpreadsheetTimeZone();
  const congregacao = c.anfitria.replace(/^congrega[cç][aã]o\s+/i, '');
  const quando = capitalizar_(dataPorExtenso_(c.data)) + (c.horario ? ', às ' + c.horario.replace(':', 'h') : '');
  const ev = eventoAgenda_(c, tz, congregacao);

  const linha = (rotulo, valor) => valor
    ? '<tr><td style="padding:4px 12px 4px 0;color:#5d6675">' + rotulo + '</td><td style="padding:4px 0"><b>' +
      escHtml_(valor) + '</b></td></tr>' : '';
  const botao = (href, texto, cor) =>
    '<a href="' + escHtml_(href) + '" style="display:inline-block;margin:6px 8px 6px 0;padding:12px 18px;' +
    'border-radius:8px;background:' + cor + ';color:#fff;text-decoration:none;font-weight:bold">' + texto + '</a>';
  const html =
    '<div style="font-family:Arial,sans-serif;max-width:560px;color:#1c2230">' +
    '<div style="background:#1e4f91;color:#fff;padding:20px;border-radius:10px 10px 0 0;text-align:center">' +
    '<div style="font-size:12px;letter-spacing:2px">CONVITE</div><div style="font-size:24px">Discurso Público</div></div>' +
    '<div style="border:1px solid #d5dde8;border-top:0;padding:20px;border-radius:0 0 10px 10px">' +
    '<p>Prezado irmão <b>' + escHtml_(c.nome) + '</b>,</p>' +
    '<p>Seu discurso público está <b>confirmado</b>. Muito obrigado pela disposição!</p>' +
    '<p style="background:#e8f0fb;padding:14px;border-radius:8px;text-align:center;font-size:18px">' +
    'Tema nº ' + c.numero + '<br><b>' + escHtml_(c.titulo) + '</b></p>' +
    '<table style="font-size:15px">' + linha('Data', quando) + linha('Congregação', congregacao) +
    linha('Endereço', c.endereco) + linha('Hospitalidade', String(r[I.HOSPITALIDADE] || '')) + '</table>' +
    '<p style="margin-bottom:0"><b>Adicione à sua agenda para ser lembrado:</b></p>' +
    '<p>' + botao(ev.linkGoogle, '📅 Google Agenda (Gmail)', '#1e4f91') +
    botao(ev.linkOutlook, '📅 Outlook / Hotmail', '#0f6cbd') + '</p><p>' +
    (c.linkConvite ? botao(c.linkConvite, 'Ver convite', '#2f6fbf') : '') +
    (c.mapa ? botao(c.mapa, '📍 Ver no mapa', '#5d6675') : '') + '</p>' +
    '<p style="font-size:13px;color:#5d6675">Usa outra agenda (iPhone, Yahoo…)? Abra o arquivo ' +
    '<b>discurso.ics</b> anexo a este e-mail.</p>' +
    '<p style="font-style:italic;color:#5d6675">Seus irmãos' +
    (congregacao ? ' da Congregação ' + escHtml_(congregacao) : '') + '</p>' +
    '</div></div>';

  MailApp.sendEmail({
    to: para,
    subject: 'Convite confirmado — Discurso Público — ' + dataCurta_(c.data),
    name: 'Discurso Público' + (congregacao ? ' — Congregação ' + congregacao : ''),
    htmlBody: html,
    body: 'Seu discurso público está confirmado: ' + quando + ' — Tema nº ' + c.numero + ' — ' + c.titulo +
      (c.endereco ? ' — ' + c.endereco : '') + '\nGoogle Agenda: ' + ev.linkGoogle +
      '\nOutlook/Hotmail: ' + ev.linkOutlook +
      (c.linkConvite ? '\nConvite: ' + c.linkConvite : ''),
    attachments: [Utilities.newBlob(ev.ics, 'text/calendar', 'discurso.ics')],
  });
  return para;
}

/** Links "Adicionar ao Google Agenda" e "Outlook/Hotmail" e arquivo .ics, com lembretes 1 semana e 1 dia antes. */
function eventoAgenda_(c, tz, congregacao) {
  const titulo = 'Discurso público' + (congregacao ? ' — Congregação ' + congregacao : '') + ' (Tema nº ' + c.numero + ')';
  const detalhes = 'Tema nº ' + c.numero + ' — ' + c.titulo + (c.linkConvite ? '\n\nConvite: ' + c.linkConvite : '');
  const local = c.endereco || congregacao;
  const utc = d => Utilities.formatDate(d, 'UTC', "yyyyMMdd'T'HHmmss'Z'");
  let datas, inicioIcs, fimIcs, outlook;
  if (/^\d{1,2}:\d{2}$/.test(c.horario)) {
    const inicio = Utilities.parseDate(c.data + ' ' + c.horario, tz, 'yyyy-MM-dd HH:mm');
    const fim = new Date(inicio.getTime() + 60 * 60 * 1000);
    const iso = d => Utilities.formatDate(d, tz, "yyyy-MM-dd'T'HH:mm:ssXXX");
    outlook = '&startdt=' + encodeURIComponent(iso(inicio)) + '&enddt=' + encodeURIComponent(iso(fim));
    datas = utc(inicio) + '/' + utc(fim);
    inicioIcs = 'DTSTART:' + utc(inicio);
    fimIcs = 'DTEND:' + utc(fim);
  } else {
    // Sem horário: evento de dia inteiro.
    const dia = c.data.replace(/-/g, '');
    const seguinte = Utilities.formatDate(new Date(Date.parse(c.data + 'T12:00:00Z') + 864e5), 'UTC', 'yyyyMMdd');
    datas = dia + '/' + seguinte;
    outlook = '&allday=true&startdt=' + c.data + '&enddt=' +
      seguinte.slice(0, 4) + '-' + seguinte.slice(4, 6) + '-' + seguinte.slice(6);
    inicioIcs = 'DTSTART;VALUE=DATE:' + dia;
    fimIcs = 'DTEND;VALUE=DATE:' + seguinte;
  }
  const linkGoogle = 'https://calendar.google.com/calendar/render?action=TEMPLATE' +
    '&text=' + encodeURIComponent(titulo) + '&dates=' + datas +
    '&details=' + encodeURIComponent(detalhes) + '&location=' + encodeURIComponent(local);
  // Outlook.com / Hotmail / Live (a agenda da Microsoft abre já com o evento preenchido).
  const linkOutlook = 'https://outlook.live.com/calendar/0/action/compose?rru=addevent' +
    '&subject=' + encodeURIComponent(titulo) + outlook +
    '&body=' + encodeURIComponent(detalhes) + '&location=' + encodeURIComponent(local);
  const icsTexto = s => String(s).replace(/([,;\\])/g, '\\$1').replace(/\n/g, '\\n');
  const alarme = gatilho => ['BEGIN:VALARM', 'ACTION:DISPLAY', 'DESCRIPTION:' + icsTexto(titulo),
    'TRIGGER:' + gatilho, 'END:VALARM'];
  const ics = [
    'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Convite Oradores//PT-BR', 'METHOD:PUBLISH',
    'BEGIN:VEVENT',
    'UID:' + Utilities.getUuid() + '@convite-oradores',
    'DTSTAMP:' + utc(new Date()),
    inicioIcs, fimIcs,
    'SUMMARY:' + icsTexto(titulo),
    'DESCRIPTION:' + icsTexto(detalhes),
    'LOCATION:' + icsTexto(local),
  ].concat(alarme('-P7D'), alarme('-P1D'), ['END:VEVENT', 'END:VCALENDAR']).join('\r\n');
  return { linkGoogle: linkGoogle, linkOutlook: linkOutlook, ics: ics };
}

/**
 * Substitui o orador de uma inscrição (pendente ou confirmada) por outro, na mesma data.
 * A antiga fica "Substituído"; a nova já nasce "Confirmado" com o tema escolhido pelo admin
 * (qualquer tema do catálogo). Devolve o id novo e a mensagem do WhatsApp para o substituto.
 */
function adminSubstituir(id, dados) {
  const usuario = exigirAdmin_();
  const nome = texto_(dados && dados.nome, 120);
  const congregacao = texto_(dados && dados.congregacao, 120);
  const numero = Number(dados && dados.tema);
  if (!nome || !congregacao) throw new Error('Informe nome e congregação do substituto.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let novoId;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const titulos = mapaTemas_(ss);
    if (!titulos[numero]) throw new Error('Escolha o tema do substituto.');
    const { aba, linha, r } = acharInscricao_(ss, id);
    const status = String(r[I.STATUS]);
    if (status !== ST_PENDENTE && status !== ST_CONFIRMADO) {
      throw new Error('Só é possível substituir inscrições pendentes ou confirmadas.');
    }
    garantirCabecalhos_(aba);
    const nomeAntigo = String(r[I.NOME]);
    aba.getRange(linha, I.STATUS + 1).setValue(ST_SUBSTITUIDO);
    aba.getRange(linha, I.OBS + 1).setValue('Substituído por ' + nome.replace(/^'/, '') + ' (' + usuario + ')');

    novoId = Utilities.getUuid().slice(0, 8).toUpperCase();
    const nova = new Array(Object.keys(I).length).fill('');
    nova[I.ID] = novoId;
    nova[I.RECEBIDO] = new Date();
    nova[I.DATA] = r[I.DATA];
    nova[I.HORARIO] = r[I.HORARIO] ? "'" + r[I.HORARIO] : '';
    nova[I.NOME] = nome;
    nova[I.CONGREGACAO] = congregacao;
    nova[I.WHATSAPP] = texto_(dados.whatsapp, 30);
    nova[I.EMAIL] = texto_(dados.email, 120);
    nova[I.GMAIL] = emailPessoal_(dados.gmail);
    nova[I.CONFIRMADO_EM] = new Date();
    nova[I.TEMAS_NUM] = "'" + numero;
    nova[I.TEMAS_TXT] = numero + ' - ' + titulos[numero];
    nova[I.TEMA_ESCOLHIDO] = numero + ' - ' + titulos[numero];
    nova[I.HOSPITALIDADE] = r[I.HOSPITALIDADE];
    nova[I.STATUS] = ST_CONFIRMADO;
    nova[I.CONFIRMADO_POR] = usuario;
    nova[I.OBS] = 'Substitui ' + nomeAntigo;
    aba.appendRow(nova);
    const n = aba.getLastRow();
    aba.getRange(n, I.RECEBIDO + 1).setNumberFormat('dd/MM/yyyy HH:mm');
    aba.getRange(n, I.DATA + 1).setNumberFormat('dd/MM/yyyy');
    aba.getRange(n, I.CONFIRMADO_EM + 1).setNumberFormat('dd/MM/yyyy HH:mm');

    // Mantém a data RESERVADA, agora em nome do substituto.
    const data = iso_(r[I.DATA], tz);
    const abaDatas = ss.getSheetByName(ABA_DATAS);
    const valores = abaDatas.getDataRange().getValues();
    const i = valores.findIndex((v, idx) => idx > 0 && iso_(v[D_DATA], tz) === data);
    if (i > 0) {
      abaDatas.getRange(i + 1, D_SITUACAO + 1, 1, 2)
        .setValues([[SIT_RESERVADA, nome.replace(/^'/, '') + ' — ' + congregacao.replace(/^'/, '')]]);
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  const mensagem = adminMensagem(novoId);
  mensagem.email = enviarEmailConvite_(SpreadsheetApp.getActiveSpreadsheet(), novoId);
  return { id: novoId, mensagem: mensagem };
}

/**
 * Cadastro manual pela administração (módulo "Novo evento"):
 *  - Discurso público designado diretamente pelo administrador (sem o orador passar pelo link):
 *    nº do tema obrigatório, entra já como Confirmado e pode receber a mensagem/e-mail de convite.
 *  - Eventos que substituem a reunião: assembleia, congresso, discurso especial, reunião especial
 *    com membro do Corpo Governante etc. Ficam no mesmo calendário e BLOQUEIAM a data na aba Datas,
 *    para nenhum orador escolhê-la pelo link.
 * Recusa a data se já houver inscrição pendente ou confirmada nela (cancele a anterior antes).
 */
function adminCriarEvento(dados) {
  const usuario = exigirAdmin_();
  dados = dados || {};
  const tipo = String(dados.tipo || '');
  if (TIPOS_EVENTO.indexOf(tipo) < 0) throw new Error('Escolha o tipo de evento.');
  const data = String(dados.data || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(data)) throw new Error('Informe a data.');
  const discurso = tipo === TIPO_DISCURSO;
  const nome = texto_(dados.nome, 120);
  const congregacao = texto_(dados.congregacao, 120);
  const numero = Number(dados.tema);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let id;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const config = lerConfig_(ss);
    const titulos = mapaTemas_(ss);
    let temaTexto;
    if (discurso) {
      if (!titulos[numero]) throw new Error('Informe um número de tema válido (1 a ' + Object.keys(titulos).length + ').');
      if (!nome || !congregacao) throw new Error('Informe o orador e a congregação.');
      temaTexto = numero + ' - ' + titulos[numero];
    } else {
      temaTexto = texto_(dados.titulo, 200) || tipo;
    }

    const abaInsc = ss.getSheetByName(ABA_INSCRICOES);
    garantirCabecalhos_(abaInsc);
    const ocupada = abaInsc.getDataRange().getValues().slice(1).find(r =>
      r[I.ID] && iso_(r[I.DATA], tz) === data &&
      [ST_PENDENTE, ST_CONFIRMADO].indexOf(String(r[I.STATUS])) >= 0);
    if (ocupada) {
      throw new Error('Já existe "' + (String(ocupada[I.NOME]) || String(ocupada[I.TEMA_ESCOLHIDO])) +
        '" (' + ocupada[I.STATUS] + ') em ' + dataCurta_(data) + '. Cancele-o na Gestão antes de cadastrar outro.');
    }

    // Aba Datas: reserva (discurso) ou bloqueia (evento); cria a linha se a data não existir.
    const abaDatas = ss.getSheetByName(ABA_DATAS);
    const valores = abaDatas.getDataRange().getValues();
    const i = valores.findIndex((v, idx) => idx > 0 && iso_(v[D_DATA], tz) === data);
    const horario = /^\d{1,2}:\d{2}$/.test(String(dados.horario || '')) ? String(dados.horario)
      : horario_(i > 0 ? valores[i][D_HORARIO] : '', tz, config);
    const situacao = discurso ? SIT_RESERVADA : 'BLOQUEADA';
    const obs = discurso ? nome + ' — ' + congregacao : temaTexto;
    const dataCelula = i > 0 ? valores[i][D_DATA] : Utilities.parseDate(data + ' 12:00', tz, 'yyyy-MM-dd HH:mm');
    if (i > 0) {
      abaDatas.getRange(i + 1, D_SITUACAO + 1, 1, 2).setValues([[situacao, obs]]);
    } else {
      abaDatas.appendRow([dataCelula, horario ? "'" + horario : '', situacao, obs]);
      abaDatas.getRange(abaDatas.getLastRow(), 1).setNumberFormat('dd/MM/yyyy');
    }

    id = Utilities.getUuid().slice(0, 8).toUpperCase();
    const linha = new Array(Object.keys(I).length).fill('');
    linha[I.ID] = id;
    linha[I.RECEBIDO] = new Date();
    linha[I.DATA] = dataCelula;
    linha[I.HORARIO] = horario ? "'" + horario : '';
    linha[I.NOME] = nome;
    linha[I.CONGREGACAO] = congregacao;
    linha[I.WHATSAPP] = texto_(dados.whatsapp, 30);
    linha[I.EMAIL] = texto_(dados.email, 120);
    linha[I.GMAIL] = emailPessoal_(dados.gmail);
    linha[I.TEMAS_NUM] = discurso ? "'" + numero : '';
    linha[I.TEMAS_TXT] = temaTexto;
    linha[I.TEMA_ESCOLHIDO] = temaTexto;
    linha[I.HOSPITALIDADE] = texto_(dados.hospitalidade, 300);
    linha[I.STATUS] = ST_CONFIRMADO;
    linha[I.CONFIRMADO_POR] = usuario;
    linha[I.CONFIRMADO_EM] = new Date();
    linha[I.OBS] = texto_(dados.obs, 300) || 'Cadastrado manualmente';
    linha[I.TIPO] = tipo;
    abaInsc.appendRow(linha);
    const n = abaInsc.getLastRow();
    abaInsc.getRange(n, I.RECEBIDO + 1).setNumberFormat('dd/MM/yyyy HH:mm');
    abaInsc.getRange(n, I.DATA + 1).setNumberFormat('dd/MM/yyyy');
    abaInsc.getRange(n, I.CONFIRMADO_EM + 1).setNumberFormat('dd/MM/yyyy HH:mm');
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  if (!discurso) return { id: id };
  const mensagem = adminMensagem(id);
  mensagem.email = enviarEmailConvite_(SpreadsheetApp.getActiveSpreadsheet(), id);
  return { id: id, mensagem: mensagem };
}

/** Marca CONVITE = Enviado (chamado quando o admin toca em "Enviar pelo WhatsApp"). */
function adminMarcarEnviado(id) {
  exigirAdmin_();
  const { aba, linha } = acharInscricao_(SpreadsheetApp.getActiveSpreadsheet(), id);
  aba.getRange(linha, I.CONVITE + 1).setValue('Enviado');
  return true;
}

/** Cancela a inscrição e devolve a data para DISPONÍVEL. */
function adminCancelar(id) {
  exigirAdmin_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const { aba, linha, r } = acharInscricao_(ss, id);
    aba.getRange(linha, I.STATUS + 1).setValue(ST_CANCELADO);
    const data = iso_(r[I.DATA], tz);
    const abaDatas = ss.getSheetByName(ABA_DATAS);
    const valores = abaDatas.getDataRange().getValues();
    const i = valores.findIndex((v, idx) => idx > 0 && iso_(v[D_DATA], tz) === data);
    const evento = String(r[I.TIPO] || TIPO_DISCURSO) !== TIPO_DISCURSO;
    if (i > 0 && (norm_(valores[i][D_SITUACAO]) === SIT_RESERVADA || (evento && norm_(valores[i][D_SITUACAO]) === 'BLOQUEADA'))) {
      abaDatas.getRange(i + 1, D_SITUACAO + 1, 1, 2).setValues([['DISPONÍVEL', '']]);
    }
    SpreadsheetApp.flush();
    return true;
  } finally {
    lock.releaseLock();
  }
}

/** Devolve o e-mail do usuário se ele for administrador; senão lança erro. */
function exigirAdmin_() {
  const email = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  const lista = String(lerConfig_(SpreadsheetApp.getActiveSpreadsheet())[CFG_ADMINS] || '')
    .toLowerCase().split(/[\s,;]+/).filter(Boolean);
  if (!email) throw new Error('Não foi possível identificar sua conta Google. Abra pelo link de administração.');
  if (!lista.includes(email)) throw new Error('A conta ' + email + ' não tem permissão de administração.');
  return email;
}

/** Cria os cabeçalhos das colunas novas (GMAIL, CONFIRMADO EM) se a planilha for da versão anterior. */
function garantirCabecalhos_(aba) {
  Object.keys(CABECALHOS_NOVOS).forEach(chave => {
    const celula = aba.getRange(1, I[chave] + 1);
    if (!celula.getValue()) celula.setValue(CABECALHOS_NOVOS[chave]);
  });
}

/**
 * E-mail pessoal do orador (Gmail, Hotmail, Outlook, Yahoo…). Recusa jwpub.org, que tem campo
 * próprio e não recebe e-mail de fora. Qualquer coisa inválida vira ''.
 */
function emailPessoal_(v) {
  const s = String(v || '').trim().toLowerCase();
  return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(s) && !/@jwpub\.org$/.test(s) ? s : '';
}

function acharInscricao_(ss, id) {
  const aba = ss.getSheetByName(ABA_INSCRICOES);
  const valores = aba.getDataRange().getValues();
  const i = valores.findIndex((r, idx) => idx > 0 && String(r[I.ID]) === String(id));
  if (i < 0) throw new Error('Inscrição ' + id + ' não encontrada.');
  return { aba: aba, linha: i + 1, r: valores[i] };
}

/** Para cada tema, a data mais recente em que já foi confirmado (ajuda a evitar repetição). */
function ultimasApresentacoes_(lista) {
  const ultimas = {};
  lista.forEach(x => {
    if (x.status === ST_CONFIRMADO && x.temaEscolhido && (!ultimas[x.temaEscolhido] || x.data > ultimas[x.temaEscolhido])) {
      ultimas[x.temaEscolhido] = x.data;
    }
  });
  return ultimas;
}

// ============================================================ utilitários

function linhas_(ss, nomeAba) {
  const aba = ss.getSheetByName(nomeAba);
  if (!aba) throw new Error('Aba "' + nomeAba + '" não encontrada na planilha.');
  return aba.getDataRange().getValues().slice(1);
}

function lerTemas_(ss) {
  return linhas_(ss, ABA_TEMAS)
    .filter(r => r[0] !== '' && String(r[1]).trim() !== '')
    .map(r => ({ numero: Number(r[0]), titulo: String(r[1]).trim() }))
    .filter(t => Number.isInteger(t.numero));
}

function mapaTemas_(ss) {
  const m = {};
  lerTemas_(ss).forEach(t => { m[t.numero] = t.titulo; });
  return m;
}

function lerConfig_(ss) {
  const cfg = {};
  if (!ss.getSheetByName(ABA_CONFIG)) return cfg;
  linhas_(ss, ABA_CONFIG).forEach(r => { cfg[norm_(r[0])] = String(r[1] || '').trim(); });
  return cfg;
}

/** Célula de data (Date ou texto dd/mm/aaaa) → 'aaaa-mm-dd'. */
function iso_(v, tz) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  const m = String(v || '').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  return m ? m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2) : '';
}

/** Horário da linha (texto "18:30" ou célula de hora) ou, se vazio, o HORÁRIO PADRÃO da Config. */
function horario_(v, tz, config) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz, 'HH:mm');
  const s = String(v || '').trim();
  return s || (config && config[CFG_HORARIO_PADRAO]) || '';
}

function numeroTema_(v) {
  const m = String(v || '').match(/^\s*(\d+)/);
  return m ? Number(m[1]) : null;
}

const MESES_ = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto',
  'setembro', 'outubro', 'novembro', 'dezembro'];
const DIAS_ = ['domingo', 'segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira', 'sexta-feira', 'sábado'];

function dataPorExtenso_(iso) {
  const [a, m, d] = iso.split('-').map(Number);
  const dt = new Date(a, m - 1, d);
  return DIAS_[dt.getDay()] + ', ' + d + ' de ' + MESES_[m - 1] + ' de ' + a;
}

function dataCurta_(iso) { return iso.split('-').reverse().join('/'); }
function capitalizar_(s) { return s.charAt(0).toUpperCase() + s.slice(1); }
function primeiroNome_(nome) { return String(nome).trim().split(/\s+/)[0]; }

/** "(85) 99999-1234" → "5585999991234" (formato do wa.me). */
function foneWhatsApp_(v) {
  let d = String(v || '').replace(/\D/g, '').replace(/^0+/, '');
  if (d.length === 10 || d.length === 11) d = '55' + d;
  return d.length >= 12 ? d : '';
}

/** Maiúsculas, sem acentos e sem espaços nas pontas — "Disponível" === "DISPONIVEL". */
function norm_(v) {
  return String(v || '').normalize('NFD').replace(/[̀-ͯ]/g, '').trim().toUpperCase();
}

/**
 * Texto do formulário: corta no tamanho máximo e neutraliza início com = + - @
 * (senão o Planilhas interpreta como fórmula — ex.: WhatsApp "+55 85 ...").
 */
function texto_(v, max) {
  const s = String(v || '').replace(/\s+/g, ' ').trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function escHtml_(s) {
  return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
