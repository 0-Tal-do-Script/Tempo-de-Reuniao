/**
 * Tempo de Reunião
 * @version 0.2.0
 * @license MIT
 * @author https://github.com/Apps-script/Tempo-de-Reuniao
 */

/** Função automática do App Script que é executada assim que o usuário abre a planilha */
function onOpen() {
  // Crio um item no menu "arquivo" para facilitar a usabilidade
  SpreadsheetApp.getUi()
    .createMenu('CALENDÁRIO')
    .addItem('Configurar (1º uso)', 'setup')
    .addItem('Importar Dados', 'importCalendar')
    .addToUi();
}

/** Função que faz a configuração inicial da aba, que basicamente é adicionar os campos de data, email e título das colunas */
function setup() {
  const sheet = SpreadsheetApp.getActiveSheet();

  // Campos de data
  sheet.getRange('A1').setValue('Data inicial');
  sheet.getRange('A2').setValue('Data final');
  const dateCells = sheet.getRange('B1:B2');
  dateCells.setBackground('#FFFBCE'); // Colorindo o fundo da célula em amarelo
  dateCells.setNumberFormat('dd/mm/yy'); // Definindo o formato das células como "data e hora"
  // Datas de exemplo
  let now = new Date();
  now.setHours(0, 0, 0, 0);
  let pastDate = new Date();
  pastDate.setMonth(pastDate.getMonth() - 3);
  pastDate.setHours(0, 0, 0, 0);
  sheet.getRange('B1').setValue(pastDate);
  sheet.getRange('B2').setValue(now);

  // Campo de email
  sheet.getRange('A3').setValue('Seu email');
  sheet.getRange('B3').setValue('bambam@example.com').setBackground('#FFFBCE');

  // Título das colunas
  sheet.getRange('A5:E5')
    .setValues([ ['Título do evento', 'Início', 'Término', 'Criado por', 'Convidados'] ])
    .setBackground('#B5BBFF');
}

/** Função que importa os dados do calendário após a configuração inicial */
function importCalendar() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const calendar = CalendarApp.getCalendarById(sheet.getRange('B3').getValue());
  const myEvents = calendar.getEvents(sheet.getRange('B1').getValue(), sheet.getRange('B2').getValue());

  // Limpando os dados da planilha antes de importar dados novos
  sheet.getRange('A6:E').clearContent();

  // Gerando a lista de eventos
  const rows = [];
  myEvents.forEach(event => {
    const guests = event.getGuestList().map(guest => guest.getEmail());

    rows.push([
      event.getTitle(),
      event.getStartTime(),
      event.getEndTime(),
      event.getCreators(),
      guests.sort().join(', '),
    ]);
  });

  // Inserindo as linhas na planilha
  sheet.getRange(6, 1, rows.length, 5).setValues(rows);
}