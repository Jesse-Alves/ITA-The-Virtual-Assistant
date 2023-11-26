#Bibliotecas
import telepot
import json
import pandas as pd

#Possiveis Utilizações
#import os
#import time

# ===================================== TAREFAS ========================================================
def buscar_nota_cobr(nota):
    acomp = pd.ExcelFile("Acomp_Equipes_Itapoan.xlsm")  # Colocar o diretorio / correto na aplicação
    BDcobr = pd.read_excel(acomp, sheet_name='Dados_Cobr')

    nota = str(nota)
    tipo = nota[0]

    # Converter coluna de Notas para String
    BDcobr['NOTA DE CORTE E RECORTE'] = BDcobr['NOTA DE CORTE E RECORTE'].astype(str)

    #Pegando dados
    status_corte = str(BDcobr.loc[BDcobr["NOTA DE CORTE E RECORTE"] == nota, "STATUS CORTE"])
    data_corte = BDcobr.loc[BDcobr["NOTA DE CORTE E RECORTE"] == nota, "CRIAÇÃO DA NOTA"]
    fim_corte = BDcobr.loc[BDcobr["NOTA DE CORTE E RECORTE"] == nota, "REALIZAÇÃO DA NOTA CORTE"]
    status_rel = str(BDcobr.loc[BDcobr["NOTA DE CORTE E RECORTE"] == nota, "STATUS REL."])
    data_rel = BDcobr.loc[BDcobr["NOTA DE CORTE E RECORTE"] == nota, "CRIAÇÃO DA NOTA REL."]
    fim_rel = BDcobr.loc[BDcobr["NOTA DE CORTE E RECORTE"] == nota, "REALIZAÇÃO DA NOTA REL."]

    print(status_corte)
    print(data_corte)
    print(fim_corte)
    print(status_rel)
    print(data_rel)
    print(fim_rel)

    #=============> Falar todas as informações da nota
    if tipo=="1":
        tipo = "CORTE"
    elif tipo =="2":
        tipo = "RECORTE"
    else:
        tipo = ""

    #Status
    if "VREL" in status_corte:
        frase = f"A última atualização desta nota de  {tipo}  está como Visitada e Realizada (VREL)."
        ITA.fala(frase)

        frase = f"A nota foi criada no dia {int(data_corte.dt.day)}/{int(data_corte.dt.month)}/{int(data_corte.dt.year)}."
        ITA.fala(frase)
        frase = f"E foi realizada no dia {int(fim_corte.dt.day)}/{int(fim_corte.dt.month)}/{int(fim_corte.dt.year)}."
        ITA.fala(frase)

        if ("VREL" in status_rel) or ("VNRE" in status_rel) or ("REDI" in status_rel):
            frase = f"E tem mais. Uma nota de religação foi gerada no dia {int(data_rel.dt.day)}/{int(data_rel.dt.month)}/{int(data_rel.dt.year)}."
            ITA.fala(frase)

        if "VREL" in status_rel:
            frase = f"E já foi realizada no dia {int(fim_rel.dt.day)}/{int(fim_rel.dt.month)}/{int(fim_rel.dt.year)}, ou seja, VREL na Religação."
            ITA.fala(frase)
        elif "VNRE" in status_rel:
            frase = f"Mas NÃO foi realizada. A tentativa da realização foi no dia {int(fim_rel.dt.day)}/{int(fim_rel.dt.month)}/{int(fim_rel.dt.year)}."
            ITA.fala(frase)
        elif "REDI" in status_rel:
            frase = "Mas a nota foi REDIRECIONADA."
            ITA.fala(frase)
        else:
            frase = "Porém não há status atualizado sobre a sua religação."
            ITA.fala(frase)

    elif "VNRE" in status_corte:
        frase = f"A última atualização desta nota de  {tipo}  está como Visitada e NÃO Realizada (VNRE)."
        ITA.fala(frase)
        frase = f"A nota foi criada no dia {int(data_corte.dt.day)}/{int(data_corte.dt.month)}/{int(data_corte.dt.year)}."
        ITA.fala(frase)
        frase = f"E foi realizada no dia {int(fim_corte.dt.day)}/{int(fim_corte.dt.month)}/{int(fim_corte.dt.year)}."
        ITA.fala(frase)
    elif "DESP" in status_corte:
        frase = f"A última atualização desta nota de {tipo} está como DESPACHADA (DESP)."
        ITA.fala(frase)
    elif "ANUL" in status_corte:
        frase = f"A última atualização desta nota de {tipo} está como ANULADA (ANUL)."
        ITA.fala(frase)
    elif "NVIS" in status_corte:
        frase = f"A última atualização desta nota de {tipo} está como NÃO VISITADA (NVIS)."
        ITA.fala(frase)
    elif "REDI" in status_corte:
        frase = f"A última atualização desta nota de {tipo} está como NÃO VISITADA (REDI)."
        ITA.fala(frase)
    else:
        frase = "Esta Nota não existe ou não existe status para ela."
        ITA.fala(frase)

    if ("DESP" in status_corte) or ("ANUL" in status_corte) or ("NVIS" in status_corte) or ("REDI" in status_corte):
        frase = f"A nota foi criada no dia {int(data_corte.dt.day)}/{int(data_corte.dt.month)}/{int(data_corte.dt.year)}."
        ITA.fala(frase)

# ================================ CLASSE DO TELEGRAM ==================================================
class Chatbot():
    def __init__(self,nome_bot):
        try:
            memoria = open('backup'+nome_bot+'.json','r')
        except FileNotFoundError:
            memoria = open('backup' + nome_bot + '.json', 'w')
            memoria.write('[["jessé"], {"o que é uma utd?": "UTD é a sigla para Unidade Territorial de Distribuição. Aqui na Coelba as regiões de atendimento são divididas nestas unidades."}]')
            memoria.close()
            memoria = open('backup' + nome_bot + '.json', 'r')
        finally:
            self.nome_bot = nome_bot
            self.conhecidos, self.conhecimento = json.load(memoria)
            memoria.close()
            self.historico = [None]
            self.usuario = ""
            self.flag_minuscula = 1

            # Frases de Conhecimento
            self.apresenta = f'Olá, tudo bom? Meu nome é {self.nome_bot}, eu sou a Assistente Virtual da UTD de ITAPOAN. Qual é o seu nome?'
    def recebendoMsg(self,msg):
        # FUNÇÃO QUE DIRECIONA A MENSAGEM PARA O BOT RACIOCINAR
        frase = self.comandos(act=msg['text'])
        #chatID = msg['chat']['id']  #Pegar somente o chatID dentro do dicionario
        self.tipoMsg, self.tipoChat, self.chatID = telepot.glance(msg)
        self.pensa(frase)
    def gravaMemoria(self):
        memoria = open('backup' + self.nome_bot + '.json', 'w')
        json.dump([self.conhecidos, self.conhecimento], memoria)
        memoria.close()
    def saudar(self,nomeUser):
        # Saudação
        if nomeUser in self.conhecidos:
            f1 = 'E aí'
            f2 = ', que bom ter você de volta'
        else:
            f1 = 'Muito prazer'
            f2 = ', eu vou gravar o seu nome aqui'
            self.conhecidos.append(nomeUser)
            self.gravaMemoria()
        nomeUser = nomeUser.title()
        self.usuario = nomeUser
        return f'{f1} {nomeUser}{f2}. Como eu posso lhe ajudar?'
    def comandos(self, act=None):
        if act == None:
            act = input('>: ')
        act = str(act)
        if self.flag_minuscula == 1:
            act = act.lower()
        return act
    def fala(self,frase):
        ita_tele.sendMessage(self.chatID,frase)
        #print(frase)
        self.historico.append(frase)
    def pensa(self,frase):
        if frase in self.conhecimento:
            resp = self.conhecimento[frase]
            self.fala(resp)
        # APRESENTAÇÃO
        elif ('oi' in frase) or ('olá' in frase) or ('e aí' in frase):
            resp = self.apresenta
            self.fala(resp)
        # GRAVAR ULTIMA MENSAGEM ENVIADA
        elif self.historico[-1] == self.apresenta:
            resp = self.saudar(frase)
            self.fala(resp)
    #================================== ENSINANDO NOVAS FUNÇÕES ==============================================
        elif ('aprende' in frase) or ('aprenda' in frase):
            self.fala('Digite a frase que será o gatilho para eu lhe responder: ')
        elif self.historico[-1] == 'Digite a frase que será o gatilho para eu lhe responder: ':
            self.chave = frase
            self.chave = str(self.chave)
            self.chave = self.chave.lower()
            self.flag_minuscula = 0
            self.fala('Agora digite o que eu devo lhe responder quando ouvir isso: ')
        elif self.historico[-1] == 'Agora digite o que eu devo lhe responder quando ouvir isso: ':
            valor = frase
            self.conhecimento[self.chave] = valor
            self.gravaMemoria()
            self.flag_minuscula = 1
            self.fala('Aprendi!! Muito obrigada!')
# =============================================== TAREFAS ====================================================
        elif ('abrir' in frase or 'preencher' in frase) and ('health check' in frase or 'healthcheck' in frase):
            self.fala(f'Excelente {self.usuario}! Você sabe que nosso lema é: Em primeiro lugar, a vida!')
            self.fala('Basta acessar este link: https://neoenergia.healthcheck.live/taking-care')
            #time.sleep(2)
            #Abrir link no windows
            #os.startfile('https://neoenergia.healthcheck.live/taking-care')
        elif ('procurar' in frase or 'buscar' in frase or 'consultar' in frase) and ('nota' in frase) and ('cobrança' in frase):
            self.fala('Pode digitar a Nota de Corte ou Recorte!')
        elif self.historico[-1] == 'Pode digitar a Nota de Corte ou Recorte!':
            buscar_nota_cobr(frase)

        # ============================================================================================================
        elif ('tchau' in frase) or ('até mais' in frase) or ('não' in frase):
            resp = f'Amei conversar contigo {self.usuario}. Até mais'
            self.fala(resp)
        else:
            try:
                resp = eval(frase)
                self.fala(resp)
            except:
                resp = 'Eu não entendi o que você falou. Posso lhe ajudar com algo mais?'
                self.fala(resp)

#Inicialização de ITA no telegram
ita_tele = telepot.Bot("2134067088:AAHBFpgjbV7MUlpYR9ULxcgTWnZSb6WFTlk")
ITA = Chatbot('ITA')

#Laço do código para conversa ficar funcionando
ita_tele.message_loop(ITA.recebendoMsg)
while True:
    pass