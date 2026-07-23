import xlwings as xw # importa a biblioteca para manipular o excel .
from datetime import date   # importa sobemente a função date da biblioteca datetime de pega a data, biblioteca do python ja.
import time # importa biblioteca para poder dar o comando de esperar 10 segundos 


app = xw.App(visible=True)   # cria a instância do Excel; visible=False roda em segundo plano
app.display_alerts = False   # suprime qualquer alerta/pop-up do Excel, incluindo esse
wb = app.books.open(
    r'U:\AREA_DE_DADOS\Indicadores\Gestao de Contratos\FILIAL SP\KPI - Mapa de Vendas\Mapa de vendas v0 - Jul26.xlsx',
    update_links=0   # 0 = não atualiza vínculos automaticamente ao abrir, e não pergunta nada
)


abrir_planilha = wb.sheets('Pedido de venda') #nessa linha tranformamos o wb.sheets(Pedido de Venda) em uma varaivel, basicamente esse sheet é uma função do wlwin
abrir_planilha.activate() #essa linha faz com que mostre pra mim que aba pedido de vendas foi aberta, pois na linha de cima ela so entrou na aba, mas não significa que mostrou pra mim, o usúario
wb.api.RefreshAll() #atualiza o arquivo todo, conexões com bancos externas também, atuaiza com os dados que recebeu antes desta linha 
time.sleep(15) #como a linha de cima so atualiza e nas proximas linhas vou precisar trabalhar com os dados atualizados, essa linha garante de uma forma bem ruim que ls dados estejam atualizados antes de eu dar o proximo comando para o excel, é uma forma que da pra melhorar, mas por enquanto, vai servir.


ultima_linha = abrir_planilha.range('P2').end('down').row #aqui range sgnifica um pedaçõ do codigo(P2) é o ponrto que ele usa como referencia, o .end(down) siginifia a mesma coisa que aperta ctrl seta ora baixo, entao vai pra ultima linha e .row te fala o nuemro dessa linha, ou seja ele usa a celula p2 como referencia, vai pra ultima linha e ega esse n8mero, oque sinigiffica a ultima linha da planilha 
ultima_coluna = abrir_planilha.range('P2').end('left').column # mesma logica da linha de cima, porem ele quer saber aultima coluna, afim de fechar o quadrado da tabela total que vai ser selecionado para ser copiado no futuro 
tabela = abrir_planilha.range((1, 1), (ultima_linha, ultima_coluna))
print(tabela)
