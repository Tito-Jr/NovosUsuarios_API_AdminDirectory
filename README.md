# Usuários google workspace
Este README fornece uma visão geral do fluxo de execução das principais funções do script desenvolvido para gerenciar usuários de um domínio google em uma plataforma educacional.

## Fluxo de Execução das Principais Script

## Funções Principais
## **1. `filtrarAlunos()`**                                                                                                                                                                                                                          

Esta função é responsável por filtrar os alunos da planilha base Siga, identificar equivalências entre os alunos e os usuários do Google Workspace e preencher as planilhas de equivalência e de novos alunos com as informações relevantes.

Passos do Fluxo:
1. Formata os nomes na planilha base Siga.
2. Obtém os alunos da planilha base Siga.
3. Itera sobre cada aluno na planilha base Siga.
   * Verifica se o e-mail do aluno está na lista de usuários por e-mail.
   * Verifica se o ID do usuário corresponde ao CPF do aluno.
   * Verifica se o aluno possui alias na lista de aliases.
   * Se nenhuma equivalência for encontrada, adiciona o aluno à lista de novos alunos.
4. Preenche as planilhas com os novos alunos e alunos com equivalências.
5. Exibe um pop-up com o número de equivalências encontradas.


## **2. `gerarEmail()`**

Esta função é responsável por gerar endereços de e-mail para os alunos com base em seus nomes.

Passos do Fluxo:
1. Obtém os nomes dos alunos da planilha de equivalência de e-mail.
2. Limpa os nomes dos alunos, removendo preposições e caracteres especiais.
3. Extrai as iniciais dos nomes para gerar o endereço de e-mail.
4. Cria o endereço de e-mail para cada aluno.
5. Preenche a coluna de e-mails na planilha de equivalência de e-mail com os endereços gerados.


## **3. `filtro2()`**

Esta função aplica um segundo filtro comparando os novos e-mails encontrados na planilha de equivalência de e-mail com os e-mails dos usuários do Google Workspace.

Passos do Fluxo:
1. Obtém os novos e-mails encontrados na planilha de equivalência de e-mail.
2. Obtém os e-mails dos usuários do Google Workspace.
3. Itera sobre cada novo e-mail encontrado.
   * Verifica se o e-mail existe na lista de e-mails dos usuários.
   * Se o e-mail existe, define as informações do usuário na coluna E da planilha de equivalência de e-mail. 


## **4. `juntarBases()`**

Esta função junta os dados das planilhas de filtragem e de equivalência de e-mail, formata os dados dos novos usuários e os adiciona à planilha 'Adicionar Usuário'.

Passos do Fluxo:
1. Obtém os novos usuários da planilha base filtrada e da planilha de equivalência de e-mail.
2. Concatena os novos usuários das duas planilhas.
3. Formata os dados dos novos usuários.
4. Preenche a planilha 'Adicionar Usuário' com os novos usuários.

   
## **5. `addUsuario()`**   

Esta função adiciona usuários ao Google Workspace com base nas informações fornecidas na lista de usuários.

Passos do Fluxo:
1. Itera sobre cada usuário na lista de usuários.
2. Preenche as informações do usuário com os dados da lista.
3. Insere o usuário no diretório do Google Workspace.
