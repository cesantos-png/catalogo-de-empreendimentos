/**
 * Script do Google Apps Script para a PLANILHA que alimenta o Catálogo de
 * Empreendimentos (a mesma cujo link "Publicar na Web" vira SHEET_CSV_URL
 * em index.html).
 *
 * Problema resolvido: hoje a coluna "Comodidades" é preenchida digitando
 * texto livre separado por vírgula, o que gera divergências de grafia
 * (acentos, plural, espaços) e cria comodidades "fantasma" no catálogo.
 * Este script adiciona um seletor de múltiplas comodidades: em vez de
 * digitar, quem preenche a planilha escolhe as opções em uma janela com
 * caixas de seleção, e o texto separado por vírgula é montado automaticamente
 * na célula — no formato que index.html já sabe interpretar.
 *
 * O Google Sheets não tem um tipo de célula "lista suspensa de múltipla
 * escolha" nativo (a validação de dados nativa só permite escolher UM
 * valor por célula), por isso a lista de múltipla escolha é implementada
 * como uma caixa de diálogo (Apps Script) em vez de uma validação de dados
 * comum.
 *
 * INSTALAÇÃO (feita uma única vez, direto na planilha do Google Sheets):
 *   1. Abra a planilha no Google Sheets.
 *   2. Menu Extensões → Apps Script.
 *   3. Apague o conteúdo do arquivo padrão (Code.gs) e cole todo o
 *      conteúdo deste arquivo no lugar.
 *   4. Salve o projeto (ícone de disquete) e recarregue a planilha.
 *   5. Um novo menu "Comodidades" aparece na planilha.
 *
 * USO:
 *   1. Clique em uma célula da coluna "Comodidades" (ou "Amenidades").
 *   2. Menu Comodidades → "Selecionar comodidades...".
 *   3. Marque as opções desejadas e clique em Salvar.
 *
 * IMPORTANTE: a lista AMENITIES abaixo deve ficar igual à constante
 * DEFAULT_AMENITIES em index.html (mesmos textos, com acento) para que os
 * ícones do catálogo combinem com o que foi selecionado na planilha. Se um
 * novo tipo de comodidade for adicionado em um dos dois lugares, adicione
 * também no outro.
 */

const AMENITY_COLUMN_ALIASES = ['comodidades', 'amenidades'];

const AMENITIES = [
  'Estacionamento',
  'Academia',
  'Piscina',
  'Carregador elétrico',
  'Pet friendly',
  'Sauna',
  'Spa',
  'Quadra esportiva',
  'Brinquedoteca',
  'Salão de jogos',
  'Copa para bebê',
  'Transfer',
  'Salão de beleza',
  'Recreação infantil',
  'Serviço de praia',
  'Sala VIP',
  'Lavanderia',
  'Loja de conveniência',
  'Coworking',
  'Bicicletário',
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Comodidades')
    .addItem('Selecionar comodidades...', 'showAmenityPicker')
    .addToUi();
}

function normalize_(text) {
  return String(text || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function findAmenityColumn_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return -1;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (let i = 0; i < headers.length; i++) {
    if (AMENITY_COLUMN_ALIASES.indexOf(normalize_(headers[i])) !== -1) {
      return i + 1; // getRange usa índice baseado em 1
    }
  }
  return -1;
}

function showAmenityPicker() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const cell = sheet.getActiveCell();

  if (cell.getRow() === 1) {
    ui.alert('Selecione uma célula de dados (não o cabeçalho) na coluna "Comodidades".');
    return;
  }

  const amenityCol = findAmenityColumn_(sheet);
  if (amenityCol === -1) {
    ui.alert('Não encontrei uma coluna "Comodidades" (ou "Amenidades") na primeira linha desta aba.');
    return;
  }

  if (cell.getColumn() !== amenityCol) {
    ui.alert('Selecione uma célula da coluna "Comodidades" antes de abrir este menu.');
    return;
  }

  const selecionadas = String(cell.getValue() || '')
    .split(/[;,]/)
    .map(function (s) { return s.trim(); })
    .filter(Boolean);

  const html = buildPickerHtml_(selecionadas, cell.getRow(), cell.getColumn());
  const output = HtmlService.createHtmlOutput(html).setWidth(360).setHeight(480);
  ui.showModalDialog(output, 'Selecionar comodidades');
}

function buildPickerHtml_(selecionadas, row, col) {
  const itens = AMENITIES.map(function (label) {
    const marcado = selecionadas.some(function (s) { return normalize_(s) === normalize_(label); }) ? 'checked' : '';
    const valor = label.replace(/"/g, '&quot;');
    return '<label style="display:block;padding:4px 0;font:14px Arial,sans-serif;">' +
      '<input type="checkbox" value="' + valor + '" ' + marcado + '> ' + label +
      '</label>';
  }).join('');

  return '<div style="font-family:Arial,sans-serif;">' +
    '<div style="max-height:340px;overflow-y:auto;border:1px solid #ddd;border-radius:6px;padding:8px 12px;">' + itens + '</div>' +
    '<div style="margin-top:14px;text-align:right;">' +
    '<button onclick="google.script.host.close()" style="margin-right:8px;">Cancelar</button>' +
    '<button onclick="salvar()" style="background:#1a73e8;color:#fff;border:none;padding:6px 14px;border-radius:4px;cursor:pointer;">Salvar</button>' +
    '</div>' +
    '<script>' +
    'function salvar(){' +
    '  var marcadas = Array.prototype.slice.call(document.querySelectorAll("input[type=checkbox]:checked")).map(function(el){ return el.value; });' +
    '  google.script.run.withSuccessHandler(function(){ google.script.host.close(); }).applyAmenitySelection(marcadas.join(", "), ' + row + ', ' + col + ');' +
    '}' +
    '</script>' +
    '</div>';
}

function applyAmenitySelection(valor, row, col) {
  SpreadsheetApp.getActiveSheet().getRange(row, col).setValue(valor);
}
