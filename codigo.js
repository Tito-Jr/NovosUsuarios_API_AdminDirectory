// Váriaveis globais
const SS = SpreadsheetApp.getActiveSpreadsheet();
const preposicoes = ['de', 'da', 'do', 'dos', 'das', 'e']; 
const caracteresEspeciais = {
  'á': 'a', 'é': 'e', 'í': 'i', 'ó': 'o', 'ú': 'u',
  'à': 'a', 'è': 'e', 'ì': 'i', 'ò': 'o', 'ù': 'u',
  'â': 'a', 'ê': 'e', 'î': 'i', 'ô': 'o', 'û': 'u',
  'ã': 'a', 'õ': 'o',
  'ç': 'c',
  'ñ': 'n'
};

const sheetBaseGoogle = SS.getSheetByName('Base Google');

const sheetBaseSiga = SS.getSheetByName('Base Siga');

const sheetAdicionarUsuarios = SS.getSheetByName('Adicionar Usuário');
var usuarios = sheetAdicionarUsuarios.getDataRange().getValues();
usuarios.shift();

const sheetBaseFiltrada = SS.getSheetByName('Base filtrada');
var novosUsuarios = sheetBaseFiltrada.getDataRange().getValues();
novosUsuarios.shift();

var sheetBaseTratada = SS.getSheetByName('Base tratada');
var alunosBaseTratada = sheetBaseFiltrada.getDataRange().getValues();
alunosBaseTratada.shift();

function criarMenu() {
  var ui = SpreadsheetApp.getUi(); 
  ui.createMenu('Gerir Usuários')
    .addItem('✅ Adicionar Usuários', 'addUsuario')
    .addSeparator()
    .addItem('💾 Baixar Base usuários', 'exportarBase')
    .addSeparator()
    .addItem('🛠️ Gerar emails', 'formatarNomes')
    .addSeparator()
    .addIntem('📌 Filtrar Alunos', 'filtrarAlunos')
    .addToUi();
}

function showPopup(mensagem){
var html = HtmlService.createHtmlOutput('<p>' + mensagem + '</p>')
.setWidth(250)
.setHeight(150);
SpreadsheetApp.getUi().showModalDialog(html, 'Script Concluído');
}

function formatarNomes() {
  var ultimaLinha = sheetBaseSiga.getLastRow();
  var nomes = sheetBaseSiga.getRange('A2:A' + ultimaLinha).getValues();

  // Função para formatar a primeira letra de cada palavra em maiúsculas
  function priMaiuscula(nome) {
    var palavras = nome.split(" ");
    var palavras_formatadas = palavras.map(function(palavra) {
      return palavra.charAt(0).toUpperCase() + palavra.slice(1);
    });
    return palavras_formatadas.join(" ");
  }

  // Função para converter preposições para minúsculas
  function converterPreposicoes(texto) {
    var palavras = texto.split(" ");
    var palavras_convertidas = palavras.map(function(palavra) {
      return preposicoes.includes(palavra.toLowerCase()) ? palavra.toLowerCase() : palavra;
    });
    return palavras_convertidas.join(" ");
  }

  // Aplicar formatação de nomes e conversão de preposições
  var nomesFormatados = [];
  for (var i = 0; i < nomes.length; i++) {
    var nome = nomes[i][0].toLowerCase(); // Converter para minúsculas
    var nomeFormatado = priMaiuscula(nome);
    nomeFormatado = converterPreposicoes(nomeFormatado);
    
    var nomesSeparados = nomeFormatado.split(" ");
    var primeiroNome = nomesSeparados.shift(); // Remover e obter o primeiro nome
    var outrosNomes = nomesSeparados.join(" "); // Obter os demais nomes
    
    nomesFormatados.push([primeiroNome, outrosNomes]);
  }

  // Escrever resultados na aba de resultados
  sheetBaseTratada.getRange('A2:B' + (nomesFormatados.length + 1)).setValues(nomesFormatados);

  // Função para remover preposições e substituir caracteres especiais
  function limparNome(nome) {
    // Substituir caracteres especiais

    var nomeLimpo = nome.replace(/[áàâãéèêíìîóòôõúùûñç]/g, function(letra) {
      return caracteresEspeciais[letra];
    });
    // Remover caracteres especiais restantes e preposições
    var palavras = nomeLimpo.split(" ");
    var palavras_sem_preposicoes = palavras.filter(function(palavra) {
      return !preposicoes.includes(palavra.toLowerCase());
    });
    return palavras_sem_preposicoes.join(" ");
  }

  function extrairIniciaisNomesMeio(nomes) {
    var iniciais = [];
    for (var i = 0; i < nomes.length; i++) {
      iniciais += nomes[i].charAt(0);
    }
    return iniciais;
  }
  
  // Função para gerar email
  function gerarEmail(nomeLimpo) {
    var partesNome = nomeLimpo.split(" ");
    var nomeUsuario = partesNome.shift();
    var sobrenome = partesNome.pop();
    var iniciaisNomesMeio = extrairIniciaisNomesMeio(partesNome);
    var email = nomeUsuario + '.'+ iniciaisNomesMeio + sobrenome + '@upe.br';
    console.log(partesNome);
    return email;
  }

  // Aplicar limpeza de nomes e geração de emails
  var emails = [];
  for (var i = 0; i < nomes.length; i++) {
    var nome = nomes[i][0].toLowerCase(); // Converter para minúsculas
    var nomeLimpo = limparNome(nome);
    var email = gerarEmail(nomeLimpo);
    emails.push([email]);
  }

  // Escrever emails na coluna C da aba 'Base tratada'
  sheetBaseTratada.getRange('C2:C' + (emails.length + 1)).setValues(emails);
}

function addUsuario() {
  var usuario = {
    "name": {
      "familyName": "",
      "givenName": ""
    },
    "primaryEmail": "",
    "recoveryEmail": "",
    "password": Math.random().toString(36),
    "orgUnitPath": "",
    "emails": [
      {
        "primary": false,
        "type": "work",
        "address": ""   // aqui vai o email secundário
      },
    ],
    "externalIds": [
      {
        "value": "", // value recebe o CPF 
        "type": "organization"
      }
    ],
    "changePasswordAtNextLogin": true
  };

  usuarios.forEach(filaUsuario => {
    usuario.name.givenName = filaUsuario[0];
    usuario.name.familyName = filaUsuario[1];
    usuario.primaryEmail = filaUsuario[2];
    usuario.orgUnitPath = filaUsuario[3];
    usuario.recoveryEmail = filaUsuario[4];
    usuario.emails[0].address = filaUsuario[5];
    usuario.externalIds[0].value = filaUsuario[6].toString();
    
    AdminDirectory.Users.insert(usuario);
  });

  console.log('Usuário criado: ');
}

function exportarBase(){
  var listaUsuarios = obterlistaUsuarios();

  sheetBaseGoogle.getRange(2,1,sheetBaseGoogle.getLastRow() - 1, 20). clearContent();
    
  sheetBaseGoogle.getRange(2, 1, listaUsuarios.length, listaUsuarios[0].length).setValues(listaUsuarios);
  
  showPopup(`${listaUsuarios.length} usuários adicionada à planilha`);
}

function obterlistaUsuarios(){

  let pageToken = '';
  let page;
  var listaUsuarios = [];
  do {
    var parametros = {
      "domain" : "etegravata.com.br",
      "orderBy": "givenName",
      "maxResults": 500,
      "pageToken": pageToken
    };
     
    page = AdminDirectory.Users.list(parametros);
    const usuariosInfo = page.users;
    if (!usuariosInfo) {
      console.log("No users found.");
      return;
    }

    for (const usuarioInfo of usuariosInfo) {
      var nome = usuarioInfo.name.fullName;
      var email = usuarioInfo.primaryEmail;
      var uniOrganizacional = usuarioInfo.orgUnitPath;
      var emailRecuperacao = usuarioInfo.recoveryEmail;
      var emailSecundario = usuarioInfo.emails[0].address;
      if (usuarioInfo.externalIds !== undefined) {
        var id = usuarioInfo.externalIds[0].value;
      }
      else {
        id = "";
      }
      listaUsuarios.push([nome, email, uniOrganizacional, emailRecuperacao, emailSecundario, id]);
    }

    pageToken = page.nextPageToken ||'';
  } while (pageToken);

  return listaUsuarios;
}
  
function listAllUsers() {
  let pageToken;
  let page;
  do {
    page = AdminDirectory.Users.list({
      domain: 'etegravata.com.br',
      orderBy: 'givenName',
      maxResults: 1,
      pageToken: pageToken
    });
    const users = page.users;
    if (!users) {
      console.log('No users found.');
      return;
    }
    // Print the user's full name and email.
    for (const user of users) {
      console.log('%s (%s)', user.name.fullName, user.primaryEmail);
    }
    pageToken = page.nextPageToken;
  } while (pageToken);
}

function filtrarAlunos() {

  const alunosBaseSiga = sheetBaseSiga.getDataRange().getValues();
  alunosBaseSiga.shift();

  const alunosBaseGoogle = sheetBaseGoogle.getDataRange().getValues();   
  alunosBaseGoogle.shift();

  let novosAlunos = [];
  var numEquivalencias = 0
  
  for (let i = 0; i < alunosBaseSiga.length; i++) { 
    const novoAluno = alunosBaseSiga[i];
    let encontrouEquivalencia = false;
    
    for (let j = 0; j < alunosBaseGoogle.length; j++) {
      const alunoAntigo = alunosBaseGoogle[j];
      
     if (novoAluno[1] === alunoAntigo[5]) { // Se houver equivalência de CPF
        sheetBaseSiga.getRange(i + 2, 10).setValue("Equivalência de CPF");
        sheetBaseSiga.getRange(i + 2, 11).setValue(j+2);
        console.log('cpf');
        numEquivalencias ++;
        encontrouEquivalencia = true;
        break;
      } else if (novoAluno[6] === alunoAntigo[1]) { // Se houver equivalência de email
        sheetBaseSiga.getRange(i + 2, 10).setValue("Equivalência de Email");
        sheetBaseSiga.getRange(i + 2, 11).setValue(j+2);
        console.log('email');
        numEquivalencias ++;
        encontrouEquivalencia = true;
        break;
      } else if (novoAluno[0] === alunoAntigo[0]) { // Se houver equivalência de nome
        sheetBaseSiga.getRange(i + 2, 10).setValue("Equivalência de Nome");
        sheetBaseSiga.getRange(i + 2, 11).setValue(j+2);
        console.log('nome');
        numEquivalencias ++;
        encontrouEquivalencia = true;
        break;
      }
    }
    
    if (!encontrouEquivalencia) {
      novosAlunos.push(novoAluno);
    }
  }
  sheetBaseFiltrada.getRange(2,1,sheetBaseFiltrada.getLastRow() - 1, 20). clearContent();
  sheetBaseFiltrada.getRange(2, 1, novosAlunos.length, novosAlunos[0].length).setValues(novosAlunos);

  showPopup(`${numEquivalencias} equivalências encontradas`);
}