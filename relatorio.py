from os import link
import main, datetime
from fpdf import FPDF

# criando o documento pdf e setando algumas configurações e funções gerais
pdf = FPDF('P','mm','A4')
pdf.add_page()
def textoNormal():
    pdf.set_text_color(0,0,0)
    pdf.set_font('Arial','',12)
def colocarEmNegrito():
    pdf.set_font('Arial','B',12)
def colocarEmSublinhado():
    pdf.set_font('Arial','U',12)
def colocarEmItalico():
    pdf.set_font('Arial','I',12)
def escreverParagrafo(paragrafo):
    textoNormal()
    pdf.multi_cell(190,5,txt=paragrafo)
    pdf.ln(h=3)

# criando um cabeçalho para o documento
def escreverChaveEValor(xPos,yPos,chave,valor):
    colocarEmNegrito()
    pdf.set_font_size(13)
    pdf.text(xPos,yPos,txt=chave)
    colocarEmSublinhado()
    pdf.text(xPos+1+len(chave)*3,yPos,txt=valor)

escreverChaveEValor(10,10,'ALUNO:','Pedro Rossi')
escreverChaveEValor(76,10,'PROFESSOR:','Fernando Silva')
escreverChaveEValor(162,10,'DATA:',str(datetime.datetime.now().day)+'/0'+str(datetime.datetime.now().month)+'/'+str(datetime.datetime.now().year))
escreverChaveEValor(10,20,'CURSO:','Jornada De Tecnologia')
escreverChaveEValor(168,20,'ESCOLA:','#kick')
pdf.line(10,23,200,23)

# definindo um título 
pdf.cell(190,15,border=0)
pdf.ln()
colocarEmNegrito()
pdf.set_font_size(13)
pdf.cell(190,10,txt='Relatório Da Taxa De Letalidade Entre Os DRS',align='C')
pdf.ln(10)

# escrevendo introdução
escreverParagrafo('O presente relatório tem como finalidade mostrar as diferentes taxas de letalidades nos Departamentos Regionais de Saúde do Estado de São Paulo, compará-las e esclarecer qual região (determinada pelas cidades pertencentes ao departamento) perdeu mais pessoas infectadas pela COVID-19 percentualmente, ou seja, qual teve a maior taxa de letalidade.')

pdf.multi_cell(190,5,txt='Para realizar o relatório foram utilizados os dados disponibilizados pelo próprio site do SEADE (Sistema Estadual de Análise de Dados) na seção sobre o coronavírus.')
pdf.ln(0.3)
colocarEmSublinhado()
pdf.set_font_size(10)
pdf.set_text_color(93,216,240)
pdf.cell(35,3,'Acesse o site do SEADE',link='https://www.seade.gov.br/coronavirus/',align='L')
pdf.ln(5)

textoNormal()
pdf.multi_cell(190,5,txt='A análise desses dados foi feita utilizando a linguagem de programação Python e as suas bibliotecas Openpyxl e Matplotlib.')
pdf.ln(0.3)

colocarEmSublinhado()
pdf.set_font_size(10)
pdf.set_text_color(93,216,240)
pdf.cell(25,3,'Site do Python',link='https://www.python.org/',align='L')
pdf.cell(3,3,'|')
pdf.cell(42,3,'Documentação Openpyxl',link='https://openpyxl.readthedocs.io/en/stable/',align='L')
pdf.cell(3,3,'|')
pdf.cell(42,3,'Documentação Matplotlib',link='https://matplotlib.org/',align='L')
pdf.ln(10)

# desenvolvimento do relatório
textoNormal()
colocarEmNegrito()
pdf.cell(190,5,txt='Contexto')
pdf.ln(5)
escreverParagrafo('Durante os anos de 2020, 2021 e 2022, o mundo sofreu com a doença COVID-19, uma doença causada pelo vírus SARS-CoV-2 que ataca o sistema respiratório do infectado.')
escreverParagrafo('A doença foi especialmente contagiosa no Brasil e um dos estados mais afetados pela doença, o primeiro a registrar um caso dela, foi o estado de São Paulo, muito devido ao fato de ser o mais populoso do país e ter maior concentração urbana.')
escreverParagrafo('O Estado de São Paulo possui dezessete Departamentos Regionais de Saúde, divisões administrativas definidas pela Secretaria de Estado da Saúde de São Paulo que foram usadas pelo SEADE durante a pandemia de COVID-19 para medir a ocupação de leitos dentro de cada departamento, isso pode ser verificado no csv de Leitos e Internaçãoes disponibilizado na seção de coronavírus do site do SEADE.')
escreverParagrafo('Este relatório usa os departamentos para medir a taxa de letalidade que cada um teve e descobrir qual departamento apresenta a maior taxa de letalidade, isto é, qual departamento perdeu mais vidas em relação ao total de infectados até o momento em que foi registrado.')
pdf.ln(5)

colocarEmNegrito()
pdf.cell(190,5,txt='Análise Dos Dados')
pdf.ln(5)
escreverParagrafo('Organizando todas as cidades em seus devidos departamentos e analisando a taxa de letalidade geral de cada departamento, foi possível chegar ao resultado do seguinte gráfico:')
pdf.image('departamentos.png', w=113, h=62, x=48.5)
escreverParagrafo('Como se pode ver o departamento com maior taxa de letalidade por COVID-19 foi o de '+main.dadosDRSLetal[0]+' com '+str(main.dadosDRSLetal[1]).replace('.',',')+'%'+' de letalidade, porém, considerando que pessoas são dados inteiros, sempre que o valor em uma casa decimal passar de 0 é necessário arredondar "para cima" nesse caso '+str(round(main.dadosDRSLetal[1])+1)+'% é a taxa de letalidade real.')
pdf.ln(5)

colocarEmNegrito()
pdf.cell(190,5,txt='Conclusão')
pdf.ln(5)
escreverParagrafo('Isso mostra que a cada 100 pessoas infectadas por COVID-19 na região do departamento de '+main.dadosDRSLetal[0]+' pelo menos '+str(round(main.dadosDRSLetal[1])+1)+' delas morrem. O resultado disso tudo, até o presente momento, é de '+main.dadosDRSLetal[2]+' pessoas mortas.')
escreverParagrafo('O departamento de '+main.dadosDRSLetal[0]+' também está um pouco acima da taxa de letalidade do Estado de São Paulo como um todo ('+str(main.mediaSP).replace('.',',')+'%) e também está acima da taxa de letalidade média do Brasil, que atualmente é de 3%.')
pdf.ln(5)

colocarEmNegrito()
pdf.cell(190,5,txt='Materiais Usados')
pdf.ln(5)

pdf.set_text_color(93,216,240)
colocarEmSublinhado()
pdf.cell(txt='Planilha com os dados.', link='https://github.com/Uns0g/relatorio-covid_python/blob/root/municipios_drs.xlsx')
pdf.ln(5)
pdf.cell(txt='Arquivo CSV baixado do SEADE.', link='https://github.com/Uns0g/relatorio-covid_python/blob/root/Dados-covid-19-municipios.csv')
pdf.ln(5)
textoNormal()
pdf.cell(txt='Bibliotecas Openpyxl e Matplotlib (links acima).')
pdf.ln(10)

pdf.set_font_size(10)
pdf.cell(75,5,txt='Para fazer esse documento em pdf foi usada a biblioteca FPDF para a linguagem Python.')
pdf.ln(3)
colocarEmSublinhado()
pdf.set_font_size(10)
pdf.set_text_color(93,216,240)
pdf.cell(40,5,txt='Clique aqui para acessar a documentação.', link='https://pyfpdf.readthedocs.io/en/latest/')

# salvando o pdf 
pdf.output('relatorio-covid_taxa-de-letalidade_pedro-rossi.pdf')