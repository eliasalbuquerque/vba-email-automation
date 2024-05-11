# Automatizador de E-mails com Anexos em Excel e PDF

Este projeto contém um conjunto de macros VBA que automatizam o envio de 
e-mails via Outlook com uma planilha Excel e um PDF como anexos, ambos gerados 
a partir de uma planilha específica no Excel. A planilha é desprotegida, 
manipulada, salva como PDF e Excel, e então protegida novamente antes do e-mail 
ser enviado.

## Instalação

1. Faça o download do arquivo .bas do repositório.
2. Abra o Microsoft Excel e a pasta de trabalho onde você deseja importar o 
   código VBA.
3. Pressione `Alt + F11` para abrir o Editor VBA.
4. No Editor VBA, vá para `Arquivo > Importar Arquivo...` e selecione o arquivo 
   .bas que você baixou. O código VBA agora deve ser visível no Editor VBA.
5. Volte para o Excel, clique em `Arquivo > Salvar Como` e escolha a opção 
   "Pasta de Trabalho Habilitada para Macro do Excel (.xlsm)" no menu suspenso 
   "Tipo". Escolha um local para salvar o arquivo e clique em "Salvar".

## Uso

Depois de instalar as macros na sua pasta de trabalho, você pode criar botões 
para executar cada macro. Aqui estão os passos para criar um botão e atribuir 
um macro a ele no Excel:

1. Vá para a aba **Developer** na faixa de opções. Se você não vê essa aba, 
   você precisará ativá-la nas opções do Excel.
2. Clique em **Insert**, e então escolha a opção **Button** sob 
   **Form Controls**.
3. Desenhe o botão na sua planilha.
4. Na caixa de diálogo que aparece, selecione o macro que você deseja atribuir 
   ao botão e clique em **OK**.

Agora, sempre que você clicar no botão, o macro será executado.