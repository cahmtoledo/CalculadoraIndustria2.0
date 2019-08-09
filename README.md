# CalculadoraIndustria2.0
A calculator of overtime and delay developed in python
--- Made by Melo
===========================================================================================

Como montar o arquivo excel:

# As regras devem ser definidas para cada funcionário. 

-Entrada I: Horário Ideal de entrada

-Saída I: Horário Ideal de saída

-Tempo de almoço: é o Tempo Ideal de Almoço

-Adicional Noturno: é o horário à partir de qual passa-se a pagar adicional noturno
(deixe vazio caso não pague adicional noturno)

-Meio-Perido: Caso o dia seja meio-periodo ele vai descontar o tempo de almoço do tempo
de hora extra. Por exemplo se você sai as 12:00, mas saiu as 17:00, no entanto almoçou
das 11:00 até as 12:00, será considerado apenas 4h de extra
obs.: se você não tiver a regra de um dia considera-se o primeiro dia da mesma semana
do dia que não seja feriado ou férias ou folga ou atestado.
Caso não exista utilizará o primeiro dia do mês que não se enquadre nas categorias citadas



# Os horários devem ser postos para cada dia de cada funcionário, se não ocorrerá um erro

-Entrada 1: O horário que o funcionário chegou na empresa

-Saida 1: O horário que o funcionário foi almoçar (caso não haja entrada 2 e saida 2
preenchidas, esse é o horário que o funcionário deixou a empresa)

-Entrada 2: Horário que o funcionário retornou do almoço

-Saída 2: Horário que o funcionário deixou a empresa

obs.: É bom informar aqui caso seja férias, folga, atestado ou feriado, mesmo essas estando
informadas nas regras. Caso haja falta informar aqui, mas deixar normal o quadro de regras


# Como usar:

-Abra o programa executável

-O programa perguntará qual planilha ele deve preencher

-Após indicar a planilha o programa lerá essa planilha indentificando os Funcionários
por começarem com Func e as regras por começarem com Regra.

-Tenha se certificado que para cada Funcionário exita uma Regra.

-Certifique-se que planilha esteja fechada

-Aperte Y ou qualquer tecla que não n para continuar

-Após cada funcionário o nome do funcionário aparecerá na tela

-Em seguida aparecerá uma mensagem de que o arquivo está salvando.

-Não feche o programa antes de ter certeza que ele salvou, caso contrário a risco de se
corromper de forma irrecuperável o arquivo

----------------------------------------------------------------------------------

Caso haja algum problema ou dúvida contate em:

(12)997503220 ou cmtoledo@usp.br
