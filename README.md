# Criar etiqueta de preço usando Pyhton
Neste projeto desenvolvi um programa com interface gráfica para criar etiquetas de preços personalizadas usando um layout em PowerPoint.
O programa irá fazer a leitura dos dados em uma planilha Google que foi publicada na Web, nessa planilha contém as informações sobre os produtos como: Descrição, preço, características, etc.
Os dados são carregados da planilha em um dataframe, depois o arquivo .pptx do layout é carregado. O pragrama identifica a quantidade de produtos na planilha e cria um slide para cada um, após isso ele vai preencher para caixa de texto no slide com os dados da planilha. Após finalizar ele salva uma cópia desse arquivo em PowerPoint com todas as alterações que estará disponível para o usuário.

No script criei uma função para cada layout de etiquetas e uma interface gráfica onde o usuário seleciona qual modelo de etiqueta deseja e clica no botão "Gerar". Fazendo isso o programa identifica quais layouts de etiquetas o usuário escolheu e chama a função correspondente para gerar as etiquetas.
Atualmente estou trabalhando com 7 layout de etiquetas que foram produzidas pelo mkt da empresa, deixei uma de exemplo no repositório.

No final da execução de cada função as informações do produto são gravadas em uma determinada aba da planilha de onde foram extraído os dados. Essas informações serão acessadas pelos usuários como uma espécie de backup para ter mais agilidade na próxima atualização de preço de um mesmo produto.

[<img src="https://i.ibb.co/x6d6qHm/etiquetas.png" alt="ETIQUETA-copia1" border="0">](https://youtu.be/y_QMUpMbgCM?si=PEvYnWi77IEdHEbf)
