// Variáveis globais
const SS = SpreadsheetApp.getActiveSpreadsheet(); // Obtém a planilha ativa
const dominio = "  "; // Domínio usado para e-mails
const preposicoes = ['de', 'da', 'do', 'dos', 'das', 'e', 'von']; // Lista de preposições para normalização

// Função para normalizar CPF
const normalizeCpf = (cpf) => { return cpf.toString().replace(/[.-]/g, '');};

// Função para normalizar nome
const normalizeNome = (str) => { return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();};

// Obter a planilha "Base Siga"
const sheetBaseSiga = SS.getSheetByName('Base Siga');

// Obter a planilha "Adicionar Usuário" e extrair os dados, removendo o cabeçalho
const sheetAdicionarUsuarios = SS.getSheetByName('Adicionar Usuário');
var usuarios = sheetAdicionarUsuarios.getDataRange().getValues();
usuarios.shift();

// Obter a planilha "Base filtrada"
const sheetBaseFiltrada = SS.getSheetByName('Base filtrada');

// Obter a planilha "Base tratada"
var sheetBaseTratada = SS.getSheetByName('Base tratada');

// Obter a planilha "Base Equi CPF" e extrair os dados, removendo o cabeçalho -> aba com as contas que possuem equivalencia de CPF
var sheetBaseEquivalenciaCPF = SS.getSheetByName('Base Equi CPF');
var equivalentesCPF = sheetBaseEquivalenciaCPF.getDataRange().getValues(); 
equivalentesCPF.shift();

// Obter a planilha "Base Equi Email" -> aba com as contas que possuem equivalencia de email
var sheetBaseEquivalenciaEmail = SS.getSheetByName('Base Equi Email');

// Função para criar o menu personalizado na interface da planilha
function criarMenu() {
  var ui = SpreadsheetApp.getUi(); // Obtém a interface de usuário da planilha
  ui.createMenu('Gerir Usuários') // Cria um menu com o nome "Gerir Usuários"
    .addItem('📌 Filtrar Alunos', 'filtrarAlunos') // Adiciona uma opção no menu para filtrar alunos
    .addSeparator() // Adiciona um separador visual entre os itens do menu
    .addItem('🛠️ Gerar emails', 'gerarEmail') // Adiciona uma opção no menu para gerar emails
    .addSeparator() // Adiciona outro separador visual
    .addItem('📌 Segundo Filtro', 'filtro2') // Adiciona uma opção no menu para executar um segundo filtro
    .addSeparator() // Adiciona outro separador visual
    .addItem('Merge', 'juntarBases') // Adiciona uma opção no menu para mesclar bases de dados
    .addSeparator() // Adiciona outro separador visual
    .addItem('✅ Adicionar Usuários', 'addUsuario') // Adiciona uma opção no menu para adicionar usuários
    .addToUi(); // Adiciona o menu à interface de usuário da planilha
}

// Função para exibir um pop-up com uma mensagem personalizada
function showPopup(mensagem) {
  // Cria uma saída HTML com a mensagem fornecida
  var html = HtmlService.createHtmlOutput('<p>' + mensagem + '</p>')
      .setWidth(250) // Define a largura do pop-up como 250 pixels
      .setHeight(150); // Define a altura do pop-up como 150 pixels
  // Exibe o pop-up modal com a saída HTML e um título "Script Concluído"
  SpreadsheetApp.getUi().showModalDialog(html, 'Script Concluído');
}

// Função para formatar nomes e converter preposições para minúsculas
function formatarNomes(sheet) {
  // Obtém o número da última linha na planilha
  var ultimaLinha = sheet.getLastRow();
  // Obtém os nomes da coluna A e os e-mails da coluna B até a última linha
  var nomes = sheet.getRange('A2:A' + ultimaLinha).getValues().flat();
  var emails = sheet.getRange('B2:B' + ultimaLinha).getValues().flat();

  // Converte todos os e-mails para minúsculas
  const emailsMinusculos = emails.map(email => email.toLowerCase());

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
    var nome = nomes[i].toLowerCase(); // Converter para minúsculas
    var nomeFormatado = priMaiuscula(nome);
    nomeFormatado = converterPreposicoes(nomeFormatado);
    
    nomesFormatados.push(nomeFormatado);
  }
  // Escrever resultados formatados de nomes e e-mails na planilha
  sheet.getRange('A2:A' + (nomesFormatados.length + 1)).setValues(nomesFormatados.map(function(nome) { return [nome]; }));
  sheet.getRange('B2:B' + (emailsMinusculos.length + 1)).setValues(emailsMinusculos.map(function(nome) { return [nome]; }));

}

// Função para gerar e-mails a partir dos nomes na planilha "Base Equi Email"
function gerarEmail(){
  // Obtém os nomes da coluna A na planilha "Base Equi Email"
  var nomes = sheetBaseEquivalenciaEmail.getRange(2, 1, sheetBaseEquivalenciaEmail.getLastRow() - 1, 1).getValues().flat();

  // Função para remover preposições e substituir caracteres especiais no nome
  function limparNome(nome) {
    var nomeLimpo = normalizeNome(nome); // Normaliza o nome removendo caracteres especiais
    var palavras = nomeLimpo.split(" "); // Divide o nome em palavras
    var palavras_sem_preposicoes = palavras.filter(function(palavra) {
      return !preposicoes.includes(palavra); // Remove preposições do nome
    });
    return palavras_sem_preposicoes.join(" "); // Junta as palavras novamente
  }

  // Função para extrair as iniciais dos nomes e sobrenomes
  function extrairIniciaisNomesMeio(nomes) {
    var iniciais = [];
    for (var i = 0; i < nomes.length; i++) {
      iniciais += nomes[i].charAt(0); // Obtém a primeira letra de cada nome
    }
    return iniciais; // Retorna as iniciais dos nomes
  }

  // Função para criar o e-mail a partir do nome limpo
  function criarEmail(nomeLimpo) {
    var partesNome = nomeLimpo.split(" "); // Divide o nome em partes
    var nomeUsuario = partesNome.shift(); // Remove o primeiro nome da lista
    var sobrenome = partesNome.pop(); // Remove o último nome da lista
    var iniciaisNomesMeio = extrairIniciaisNomesMeio(partesNome); // Obtém as iniciais dos nomes do meio
    var email = nomeUsuario + '.'+ iniciaisNomesMeio + sobrenome + '@upe.br'; // Cria o e-mail
    return email; // Retorna o e-mail gerado
  }

  // Aplicar limpeza de nomes e geração de e-mails
  var emails = [];
  for (var i = 0; i < nomes.length; i++) {
    var nome = nomes[i];
    var nomeLimpo = limparNome(nome); // Limpa o nome
    var email = criarEmail(nomeLimpo); // Cria o e-mail
    emails.push([email]); // Adiciona o e-mail à lista de e-mails
  }
  // Escrever e-mails na coluna B da planilha 'Base Equi Email'
  sheetBaseEquivalenciaEmail.getRange('B2:B' + (emails.length + 1)).setValues(emails.map(function(email) { return [email]; }));
}

// Função para adicionar usuários ao Google Workspace
function addUsuario() {
  // Definição do modelo de usuário
  var usuario = {
    "name": {
      "familyName": "",   // Sobrenome do usuário
      "givenName": ""     // Nome do usuário
    },
    "primaryEmail": "",   // E-mail principal do usuário
    "recoveryEmail": "",  // E-mail de recuperação do usuário
    "password": "",       // Senha do usuário
    "orgUnitPath": "",    // Caminho da unidade organizacional do usuário
    "emails": [
      {
        "primary": false,
        "type": "work",
        "address": ""   // Endereço de e-mail secundário do usuário
      },
    ],
    "externalIds": [
      {
        "value": "", // CPF ou Matricula do usuário
        "type": "organization"
      }
    ],
    "changePasswordAtNextLogin": true // O usuário será obrigado a alterar a senha no próximo login
  };

  // Itera sobre cada usuário na lista 'usuarios'
  usuarios.forEach(filaUsuario => {
    // Preenche as informações do usuário com os dados da lista
    usuario.name.givenName = filaUsuario[0];       // Define o nome do usuário
    usuario.name.familyName = filaUsuario[1];      // Define o sobrenome do usuário
    usuario.primaryEmail = filaUsuario[2];         // Define o e-mail principal do usuário
    usuario.orgUnitPath = filaUsuario[3];          // Define o caminho da unidade organizacional do usuário
    usuario.recoveryEmail = filaUsuario[4];        // Define o e-mail de recuperação do usuário
    usuario.emails[0].address = filaUsuario[4];    // Define o endereço de e-mail secundário do usuário (igual ao e-mail de recuperação)
    usuario.password = filaUsuario[5].toString();  // Define a senha do usuário (convertendo para string)
    usuario.externalIds[0].value = filaUsuario[5].toString(); // Define o CPF do usuário (convertendo para string)
    
    // Insere o usuário no diretório do Google Workspace
    AdminDirectory.Users.insert(usuario);
  });

  console.log('Usuário criado: '); // Loga a mensagem "Usuário criado" no console
}

// Função para obter a lista de usuários do Google Workspace
function obterlistaUsuarios() {
  let pageToken = ''; // Token da próxima página de resultados
  let page; // Variável para armazenar a página de resultados
  var lista1 = {}; // Lista de usuários indexada por e-mail
  var lista2 = {}; // Lista de usuários indexada por ID (CPF ou matrícula)
  let listaAlias1 = []; // Lista de aliases não formatados
  let listaAlias2 = []; // Lista de aliases formatados

  // Loop para percorrer todas as páginas de resultados
  do {
    // Parâmetros da consulta para listar usuários
    var parametros = {
      "domain": dominio, // Domínio dos usuários
      "orderBy": "givenName", // Ordenar por nome
      "maxResults": 500, // Máximo de resultados por página
      "pageToken": pageToken // Token da página atual
    };

    // Obtém a página de resultados de usuários
    page = AdminDirectory.Users.list(parametros);
    const usuariosInfo = page.users; // Informações dos usuários na página

    // Loop para processar informações de cada usuário na página
    for (const usuarioInfo of usuariosInfo) {
      var nome = usuarioInfo.name.fullName; // Nome completo do usuário
      var email = usuarioInfo.primaryEmail; // E-mail principal do usuário
      var id; // ID do usuário (CPF ou matrícula)

      // Verifica se o usuário possui um ID externo (CPF ou matrícula)
      if(usuarioInfo.externalIds !== undefined){
        id = normalizeCpf(usuarioInfo.externalIds[0].value); // Normaliza o ID (CPF). Remove '.' e '-' 
        lista2[id] = {nome : nome, email : email}; // Adiciona o usuário à lista indexada por ID
      }
      else{
        id = undefined; // Define o ID como undefined caso não exista
      }
      lista1[email] = { nome: nome, id: id }; // Adiciona o usuário à lista indexada por e-mail

      // Verifica se o usuário possui aliases
      if(usuarioInfo.aliases){
        usuarioInfo.aliases.unshift(usuarioInfo.name.fullName); // Adiciona o nome completo como primeiro alias
        listaAlias1.push(usuarioInfo.aliases); // Adiciona a lista de aliases não formatados
      }
    }
    // Formata a lista de aliases
    listaAlias1.forEach(usuario => {
      let nome = usuario[0]; // Nome do usuário
      let emails = usuario.slice(1); // Remove o primeiro elemento (nome) do array
      // Para cada email, cria uma nova entrada na lista de aliases formatados
      emails.forEach(email => {
        listaAlias2.push([nome, email]);
      });
    });

    pageToken = page.nextPageToken || ''; // Atualiza o token da próxima página ou define como vazio se não houver mais páginas
  } while (pageToken); // Continua o loop enquanto houver uma próxima página de resultados
  
  // Retorna as listas de usuários por e-mail, por ID e a lista de aliases formatados
  return { listaByEmail: lista1, listaById: lista2, listaAlias: listaAlias2 };  
}

// Função para filtrar alunos e identificar equivalências entre a base de dados e os usuários do Google Workspace
function filtrarAlunos() {
  let novosAlunos = []; // Array para armazenar novos alunos não encontrados na base de dados do Google Workspace
  let alunosEquiCPF = []; // Array para armazenar alunos com equivalência de CPF
  let alunosEquiEmail = []; // Array para armazenar alunos com equivalência de e-mail

  // Obtém as listas de usuários do Google Workspace e aliases
  var listas = obterlistaUsuarios();
  const listaUsuariosByemail = listas.listaByEmail;
  const listaUsuariosByid = listas.listaById;
  const listaAlias = listas.listaAlias;

  var numEquivalencias = 0; // Contador de equivalências encontradas

  // Formata os nomes na planilha base Siga
  formatarNomes(sheetBaseSiga);
  var alunosBaseSiga = sheetBaseSiga.getRange(2, 1, sheetBaseSiga.getLastRow() - 1, 4).getValues(); // Obtém os dados dos alunos da planilha base Siga

  // Loop para percorrer cada aluno na planilha base Siga
  for (let i = 0; i < alunosBaseSiga.length; i++) { 
    const novoAluno = alunosBaseSiga[i]; // Dados do novo aluno
    let encontrouEquivalencia = false; // Flag para indicar se foi encontrada alguma equivalência para o aluno

    // Verifica se o e-mail do novo aluno está na lista de usuários por e-mail
    if (listaUsuariosByemail[novoAluno[1]]) {
      // Verifica se o ID do usuário corresponde ao CPF do novo aluno
      if (listaUsuariosByemail[novoAluno[1]].id == normalizeCpf(novoAluno[2])){
        sheetBaseSiga.getRange(i + 2, 5).setValue("Equivalência de Email e CPF."); // Marca a equivalência na planilha
        sheetBaseSiga.getRange(i + 2, 6).setValue(listaUsuariosByemail[novoAluno[1]]); // Exibe informações do usuário na planilha
        alunosEquiCPF.push(novoAluno); // Adiciona o aluno à lista de alunos com equivalência de CPF
        numEquivalencias ++; // Incrementa o contador de equivalências
        encontrouEquivalencia = true; // Define a flag de equivalência como verdadeira
      } else{
        sheetBaseSiga.getRange(i + 2, 5).setValue("Equivalência de Email."); // Marca a equivalência na planilha
        sheetBaseSiga.getRange(i + 2, 6).setValue(listaUsuariosByemail[novoAluno[1]]); // Exibe informações do usuário na planilha
        alunosEquiEmail.push(novoAluno); // Adiciona o aluno à lista de alunos com equivalência de e-mail
        numEquivalencias ++; // Incrementa o contador de equivalências
        encontrouEquivalencia = true; // Define a flag de equivalência como verdadeira
      }
    } else if(listaUsuariosByid[normalizeCpf(novoAluno[2])]){
        sheetBaseSiga.getRange(i + 2,5).setValue("Equivalência de CPF."); // Marca a equivalência na planilha
        sheetBaseSiga.getRange(i + 2, 6).setValue(listaUsuariosByid[normalizeCpf(novoAluno[2])]); // Exibe informações do usuário na planilha
        novoAluno[1] = listaUsuariosByid[normalizeCpf(novoAluno[2])].email; // Atualiza o e-mail do aluno com o e-mail encontrado
        alunosEquiCPF.push(novoAluno); // Adiciona o aluno à lista de alunos com equivalência de CPF
        numEquivalencias ++; // Incrementa o contador de equivalências
        encontrouEquivalencia = true; // Define a flag de equivalência como verdadeira
    }

    // Verifica se o aluno possui alias na lista de aliases
    for(let j = 0; j < listaAlias.length; j++ ){
      const alias = listaAlias[j][1];
      if (alias === novoAluno[1]) {
        sheetBaseSiga.getRange(i + 2,5).setValue("Equivalência de Alias."); // Marca a equivalência na planilha
        sheetBaseSiga.getRange(i + 2, 6).setValue(listaAlias[j]); // Exibe informações do alias na planilha
        alunosEquiEmail.push(novoAluno); // Adiciona o aluno à lista de alunos com equivalência de e-mail
        numEquivalencias ++; // Incrementa o contador de equivalências
        encontrouEquivalencia = true; // Define a flag de equivalência como verdadeira
      }
    }

    // Se nenhuma equivalência foi encontrada, adiciona o aluno à lista de novos alunos
    if (!encontrouEquivalencia) {
      novosAlunos.push(novoAluno);
    }
  }

  // Preenche as planilhas com os novos alunos e alunos com equivalências
  preencher(sheetBaseFiltrada, novosAlunos); // Preenche a planilha base filtrada com os novos alunos
  preencher(sheetBaseEquivalenciaCPF, alunosEquiCPF); // Preenche a planilha base de equivalência de CPF com os alunos com equivalência de CPF
  preencher(sheetBaseEquivalenciaEmail, alunosEquiEmail); // Preenche a planilha base de equivalência de e-mail com os alunos com equivalência de e-mail

  // Exibe um pop-up com o número de equivalências encontradas
  showPopup(`${numEquivalencias} equivalências encontradas`);
} 

// Função para preencher uma planilha do Google Sheets com dados de alunos
function preencher(sheet, alunos) {
  // Verifica se a planilha já possui dados a partir da segunda linha
  if (sheet.getLastRow() >= 2) {
    // Limpa o conteúdo da planilha, exceto pela primeira linha (cabeçalho)
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 20).clearContent();
  }
  
  // Preenche a planilha com os dados dos alunos
  sheet.getRange(2, 1, alunos.length, alunos[0].length).setValues(alunos);
}

// Função para aplicar um segundo filtro comparando os novos e-mails encontrados na planilha de equivalência de e-mail com os e-mails dos usuários do Google Workspace
function filtro2() {
  // Obtém os novos e-mails encontrados na planilha de equivalência de e-mail
  const novosEmail = sheetBaseEquivalenciaEmail.getRange('B2:B' + sheetBaseEquivalenciaEmail.getLastRow()).getValues().flat();
  // Obtém os e-mails dos usuários do Google Workspace
  const emails = obterlistaUsuarios().listaByEmail;

  var cont = 0; // Contador de linhas na planilha de equivalência de e-mail
  
  // Itera sobre cada novo e-mail encontrado na planilha de equivalência de e-mail
  novosEmail.forEach(function(email) {
    // Verifica se o e-mail existe na lista de e-mails dos usuários do Google Workspace
    if (emails[email]) {
      // Se o e-mail existe, define as informações do usuário na coluna E da planilha de equivalência de e-mail
      sheetBaseEquivalenciaEmail.getRange(cont + 2, 5).setValue(emails[email]);
    }
    cont++; // Incrementa o contador de linhas
  });
}

// Função para juntar dados das planilhas de filtragem e equivalência de e-mail e adicionar à planilha 'Adicionar Usuário'
function juntarBases() {
  // Obtém os novos usuários da planilha base filtrada
  var novosUsuarios = sheetBaseFiltrada.getRange(2, 1, sheetBaseFiltrada.getLastRow() - 1, 4).getValues();
  // Obtém os novos usuários da planilha de equivalência de e-mail
  var novosUsuarios2 = sheetBaseEquivalenciaEmail.getRange(2, 1, sheetBaseEquivalenciaEmail.getLastRow() - 1, 4).getValues();
  // Concatena os novos usuários das duas planilhas
  var novosUsuarios3 = novosUsuarios.concat(novosUsuarios2);

  var addAluno = []; // Array para armazenar os novos usuários formatados para adicionar à planilha 'Adicionar Usuário'
  
  // Itera sobre cada novo usuário
  novosUsuarios3.forEach(function(aluno) {
    // Separa o nome em partes
    var nomeSeparado = aluno[0].split(" ");
    var firstName = nomeSeparado.shift(); // Obtém o primeiro nome
    var lastName = nomeSeparado.join(" "); // Obtém o sobrenome
    var primaryEmail = aluno[1]; // Obtém o e-mail principal
    var id = aluno[2]; // Obtém o ID (CPF ou outro identificador)
    var email2 = aluno[3]; // Obtém o e-mail secundário, se existir

    // Adiciona o novo usuário formatado ao array 'addAluno'
    addAluno.push([firstName, lastName, primaryEmail, "unidade", email2, id]);
  });

  // Preenche a planilha 'Adicionar Usuário' com os novos usuários
  preencher(sheetAdicionarUsuarios, addAluno); 
}










