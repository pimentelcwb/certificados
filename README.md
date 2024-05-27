# Automação de Certificados com VBA
Este projeto automatiza a criação de certificados personalizados a partir de uma lista de nomes no Excel. Utiliza VBA para ler os nomes de uma planilha do Excel, inserir cada nome em um template de certificado no PowerPoint e exportar cada certificado como um arquivo PDF.

## Requisitos
Microsoft Excel
Microsoft PowerPoint

## Configuração
1. Preparar o Template do Certificado no PowerPoint:
- Crie um slide no PowerPoint que servirá como template para o certificado.
- Insira uma caixa de texto onde o nome deve aparecer e use um placeholder, por exemplo, "[NOME]".

2. Preparar a Lista de Nomes no Excel:
- Crie uma planilha no Excel e insira a lista de nomes na coluna A (uma linha para cada nome).

3. Adicionar o Script VBA no Excel:
- Pressione ALT + F11 para abrir o Editor do VBA.
- Insira um novo módulo (Menu Inserir > Módulo).
- Copie e cole o código de certificados_vba.

## Uso
1. Salve o arquivo Excel como uma Pasta de Trabalho Habilitada para Macro (.xlsm).
2. Pressione ALT + F8 para abrir a janela de macros no Excel.
3. Selecione CriarCertificados e clique em "Executar".

## Notas
- Substitua "C:\caminho\para\seu\template.pptx" pelo caminho completo do seu arquivo de template do PowerPoint.
- Substitua "C:\caminho\para\salvar\" pelo caminho onde você deseja salvar os PDFs.
- O script assume que o template tem apenas um slide, que é o slide do certificado.
