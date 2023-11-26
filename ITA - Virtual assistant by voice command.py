# ITA - Assitente Virtual da UTD de Itapoan
# Auto: Jessé Alves
# print(sr.Microphone.list_microphone_names())  # Reconhecer o Microfone a ser utilizado


#Bibliotecas
import speech_recognition as sr
import pyttsx3
from pygame import mixer
from tkinter import *
#import time
import pandas as pd


# ============================================================= FUNÇÕES ======================================================================================
# Ler comando de Voz
def comandos():
    comando_de_voz = ""
    try:
        rec = sr.Recognizer()  # depois ver se esse comando precisa ficar aqui dentro
        with sr.Microphone(1) as mic:
            rec.adjust_for_ambient_noise(mic)

            # Som de beep para indicar momento de falar
            mixer.init()
            mixer.music.load('C:/Users/Jesse Alves/PycharmProjects/Conversa_ITA/som_beep.mp3')
            mixer.music.play()
            #print("Fale")

            audio = rec.listen(mic)
            comando_de_voz = rec.recognize_google(audio, language="pt-BR")
    except:
        frase = "Eu não te escutei.  Pode falar de novo?"
        resp(frase)
        comando_de_voz = comandos()
    finally:
        return comando_de_voz

# Resposta de ITA
def resp(frase):
    #Inicializando a voz de ITA
    ITA = pyttsx3.init()
    ITA.say(frase)
    ITA.runAndWait()
    #Tela carregando

# ============================================================= TAREFAS ================================================================================
def buscar_nota_cobr(nota):
    acomp = pd.ExcelFile("Acomp_Equipes_Itapoan.xlsm")  # Coloca o diretorio /
    BDcobr = pd.read_excel(acomp, sheet_name='Dados_Cobr')

    sair = 0
    while sair==0:
        print("Nota sem tratar")
        print(nota)

        nota = str(nota)
        nota.replace(" ","")
        nota.replace("/","")
        tipo = nota[0]
        #nota = "2004916665" #Manualmente
        nota = nota[len(nota) - 4:len(nota)] #Pegando somente os 4 ultimos elementos
        print("Nota tratada")
        print(nota)


        # Converter coluna de Notas para String
        BDcobr['NOTA DE CORTE E RECORTE'] = BDcobr['NOTA DE CORTE E RECORTE'].astype(str)
        # Pegando os quatro ultimos elementos da string
        BDcobr['NOTA DE CORTE E RECORTE'] = BDcobr['NOTA DE CORTE E RECORTE'].str[6:10]

        #Pegando dados
        status_corte = str(BDcobr.loc[BDcobr["NOTA DE CORTE E RECORTE"] == nota, "STATUS CORTE"])
        data_corte = BDcobr.loc[BDcobr["NOTA DE CORTE E RECORTE"] == nota, "CRIAÇÃO DA NOTA"]
        fim_corte = BDcobr.loc[BDcobr["NOTA DE CORTE E RECORTE"] == nota, "REALIZAÇÃO DA NOTA CORTE"]
        status_rel = str(BDcobr.loc[BDcobr["NOTA DE CORTE E RECORTE"] == nota, "STATUS REL."])
        data_rel = BDcobr.loc[BDcobr["NOTA DE CORTE E RECORTE"] == nota, "CRIAÇÃO DA NOTA REL."]
        fim_rel = BDcobr.loc[BDcobr["NOTA DE CORTE E RECORTE"] == nota, "REALIZAÇÃO DA NOTA REL."]


        #print(status_corte)
        #print(data_corte)
        #print(fim_corte)
        #print(status_rel)
        #print(data_rel)
        #print(fim_rel)

        #=============> Falar todas as informações da nota
        if tipo=="1":
            tipo = "corte"
        elif tipo =="2":
            tipo = "recorte"
        else:
            tipo = ""

        #Status
        if "VREL" in status_corte:
            frase = f"A última atualização desta nota de  {tipo}  está como visitada e realizada."
            resp(frase)
            frase = f"A nota foi criada no dia {int(data_corte.dt.day)}.  do {int(data_corte.dt.month)}.  de " \
                    f"{int(data_corte.dt.year)}. E foi realizada no dia {int(fim_corte.dt.day)}.  do " \
                    f"{int(fim_corte.dt.month)}.  de {int(fim_corte.dt.year)}"
            resp(frase)

            if "VREL" in status_rel:
                frase = f"E tem mais.  Uma nota de religação foi gerada no dia {int(data_rel.dt.day)}.  do" \
                        f" {int(data_rel.dt.month)}.  de {int(data_rel.dt.year)}.  E já foi realizada no dia. " \
                        f"{int(fim_rel.dt.day)}.  do {int(fim_rel.dt.month)}.  de {int(fim_rel.dt.year)}"
                resp(frase)
            elif "VNRE" in status_rel:
                frase = f"E tem mais.  Uma nota de religação foi gerada no dia {int(data_rel.dt.day)}.  do" \
                        f" {int(data_rel.dt.month)}.  de {int(data_rel.dt.year)}, mas não foi realizada. A tentativa da realização foi" \
                        f" no dia {int(fim_rel.dt.day)}.  do {int(fim_rel.dt.month)}.  de {int(fim_rel.dt.year)}"
                resp(frase)
            elif "REDI" in status_rel:
                frase = f"E tem mais.  Uma nota de religação foi gerada no dia {int(data_rel.dt.day)} do" \
                        f" {int(data_rel.dt.month)} de {int(data_rel.dt.year)}, mas a nota foi redirecionada."
                resp(frase)
            else:
                frase = "Porém não há status atualizado sobre a sua religação."
                resp(frase)

        elif "VNRE" in status_corte:
            frase = f"A última atualização desta nota de  {tipo}  está como visitada e não realizada."
            resp(frase)
            frase = f"A nota foi criada no dia  {int(data_corte.dt.day)}.  do {int(data_corte.dt.month)}. de  {int(data_corte.dt.year)}. " \
                    f" E tentou ser realizada no dia  {int(fim_corte.dt.day)}.  do {int(fim_corte.dt.month)}. de  {int(fim_corte.dt.year)}"
            resp(frase)
        elif "DESP" in status_corte:
            frase = f"A última atualização desta nota de {tipo} está como despachada. A nota foi criada no dia " \
                    f"{int(data_corte.dt.day)}.  do {int(data_corte.dt.month)}.  de {int(data_corte.dt.year)}."
            resp(frase)
        elif "ANUL" in status_corte:
            frase = f"A última atualização desta nota de {tipo} está como anulada.  A nota foi criada no dia " \
                    f"{int(data_corte.dt.day)}.  do {int(data_corte.dt.month)}.  de {int(data_corte.dt.year)}."
            resp(frase)
        elif "NVIS" in status_corte:
            frase = f"A última atualização desta nota de {tipo} está como não visitada.  A nota foi criada no dia " \
                    f"{int(data_corte.dt.day)}.  do {int(data_corte.dt.month)}.  de {int(data_corte.dt.year)}."
            resp(frase)
        elif "REDI" in status_corte:
            frase = f"A última atualização desta nota de {tipo} está como redirecionada.  A nota foi criada no dia " \
                    f"{int(data_corte.dt.day)}.  do {int(data_corte.dt.month)}.  de {int(data_corte.dt.year)}."
            resp(frase)
        else:
            frase = "Não existe status atual para essa nota"
            resp(frase)


        #Verificar se deseja consultar outras notas
        frase = "Deseja consultar outra nota de cobrança?"
        resp(frase)
        fala_fim = comandos()
        if ("Sim" in fala_fim) or ("sim" in fala_fim) or ("Claro" in fala_fim) or ("claro" in fala_fim) or (
                "por favor" in fala_fim) or ("Por favor" in fala_fim) or ("Positivo" in fala_fim) or (
                "positivo" in fala_fim):

            frase = "Pode falar a próxima nota pausadamente"
            resp(frase)
            nota = comandos()
        else:
            frase = "Ok"
            resp(frase)
            sair = 1

# ========================================================== FUNÇÃO DA CONVERSA =============================================================================

# Corpo da Conversa
def conversa(tela_inicial,fr_conversa):
    fr_conversa.tkraise()

    # ====================== FLAGS ==================================
    # Flag que determinará o fim do código
    flag_fim = 0
    # Se o usuário não passar o nome dele
    nome_user = ""
    # ===============================================================

    #Chamar reconhecimento de voz
    act = comandos()
    print(act)

    # ============================== SAUDAÇÃO ANTES DE PERGUNTAR O QUE FAZER =================================
    if ("Olá" in act) or ("Oi" in act) or ("E aí" in act):
        # Frase de Saudação
        saudacao = "Olá.  Tudo bom?   Meu nome é ITA!   Eu sou a Assistente virtual da u tê dê de Itapoan. Qual é o seu nome?"
        resp(saudacao)

        #Pegou o nome do usuario
        nome_user = comandos()

        #Pergunta primeira ação
        frase = f"Muito prazer {nome_user}.  Como eu posso te ajudar?"
        resp(frase)
        act = comandos()
    #=========================================================================================================


    # ========================================== INICIAR A CONVERSA EM UM LAÇO DE REPETIÇÃO ================================================================

    # Laço de repetição da conversa
    while flag_fim == 0:
        # ============================================== Lista de todas as ações que ela pode executar =======================================
        if "formulário" in act:
            frase = "Qual formulário você quer preencher?"
            resp(frase)
            tipo_form = comandos()
            if ("Health" in tipo_form) or ("Check" in tipo_form) or ("rack" in tipo_form):
                frase = "Só um momento que eu vou abrir para você!"
                resp(frase)
                #print(tipo_form)
            else:
                frase = "Não consigo abrir esse formulário para você!"
                resp(frase)
        elif (("status" in act) or ("situação" in act) or ("consultar" in act) or ("Consultar" in act) or ("consulta" in act)) and ("nota" in act or "Nota" in act) and ("cobrança" in act):
            frase = "Eu vou consultar as informações da nota para você.  Basta me " \
                    "dizer a nota número por número pausadamente"
            resp(frase)
            nota = comandos()

            buscar_nota_cobr(nota)

        # ======================================================= Lista de Perguntas Pessoais ================================================
        elif ("namora" in act) or ("namorado" in act) or ("namorando" in act) or (("saindo" in act) and ("alguém" in act)):
            frase = f"Ái  ái {nome_user}.  Tomara que Luciano não escute nossa conversa.   Mas o Ian da Neoenergia " \
                    f"me chamou para sair essa sexta a noite. Não conta para ninguém, e vamos voltar para os assuntos " \
                    f"da empresa!"
            resp(frase)
        elif ("sobre" in act) and ("você" in act):
            frase = "A ideia de criar uma assistente virtual veio de Luciano Coelho.  Então ele procurou o melhor " \
                    "estagiário da Coelba,  chamado Jessé, e pediu que ele me desenvolvêsse.  Aos poucos estou " \
                    "ficando cada vez mais inteligente e capaz de ajudar nos processos da u tê dê de Itapoan."
            resp(frase)
        elif ("melhor" in act) and ("supervisor" in act):
            frase = "De acordo com minha busca em toda a Iberdrola, Luciano Coelho está sendo o melhor Supervisor."
            resp(frase)
        # ============================================ EM CASO DE NÃO ENTENDER OU NÃO EXISITR O COMANDO ======================================
        else:
            frase = "Eu não consegui entender o que você disse."
            resp(frase)

        # ================================================ VERIFICAR SE A CONVERSA CHEGOU AO FIM ===============================================
        frase = "Posso te ajudar com algo mais?"
        resp(frase)
        fala_fim = comandos()
        # print(fala_fim)
        if ("Sim" in fala_fim) or ("sim" in fala_fim) or ("Claro" in fala_fim) or ("claro" in fala_fim) or (
                "por favor" in fala_fim) or ("Por favor" in fala_fim) or ("Positivo" in fala_fim) or (
                "positivo" in fala_fim):
            frase = "Legal!  Me diga o próximo comando!"
            resp(frase)
            act = comandos()
        else:
            frase = "Eu entendi que não."
            resp(frase)
            flag_fim = 1


    # ============================================================ FIM DO CÓDIGO =======================================================================
    #Tela de Tchau

    # Mensagem de Tchau
    if nome_user == "":
        fala_tchau = "Foi um prazer conversar com você!  Até a próxima!  Abraços"
        resp(fala_tchau)
    else:
        fala_tchau = f"Amei conversar contigo {nome_user}!  Até a próxima!  Abraços"
        resp(fala_tchau)

    tela_inicial.destroy()



#============================================================== INICIANDO LAYOUT ===========================================================================

tela_inicial = Tk()
tela_inicial.title("Conversar com ITA")
tela_inicial.geometry("490x560+610+153")
#master.iconbitmap(default="icones\\ico.ico")
tela_inicial.resizable(width=1,height=1)


#Importar imagens
img_inicial = PhotoImage(file="imagens\\Fundo_Layout.png")
img_botao = PhotoImage(file="imagens\\botao_falar.png")
img_conversa = PhotoImage(file="imagens\\tela_conversa.png")
#img_carreg = PhotoImage(file="imagens\\tela_carregando.png")
#img_ITAfalando = PhotoImage(file="imagens\\tela_ITAfalando.png")
#img_ouvindo = PhotoImage(file="imagens\\tela_ouvindo.png")

# ====> Criação de Frames
fr_inicial = Frame(tela_inicial,borderwidth=1,relief="solid")
fr_inicial.place(x=0,y=0,width=490,height=560)

fr_conversa = Frame(tela_inicial,borderwidth=1,relief="solid")
fr_conversa.place(x=0,y=0,width=490,height=560)

#fr_carreg = Frame(tela_inicial,borderwidth=1,relief="solid")
#fr_ITAfalando = Frame(tela_inicial,borderwidth=1,relief="solid")
#fr_ouvindo = Frame(tela_inicial,borderwidth=1,relief="solid")

# Ajustando o tamanho dos Frames
#for telas in (fr_inicial, fr_carreg, fr_ITAfalando, fr_ouvindo):
    #telas.place(x=0,y=0,width=490,height=560)

#Criação de labels
lb_inicial = Label(fr_inicial,image=img_inicial)
lb_inicial.pack()

lb_conversa = Label(fr_conversa,image=img_conversa)
lb_conversa.pack()

#lb_carreg = Label(fr_carreg, image=img_carreg)
#lb_carreg.pack()

#lb_ITAfalando = Label(fr_ITAfalando, image=img_ITAfalando)
#lb_ITAfalando.pack()

#lb_ouvindo = Label(fr_ouvindo, image=img_ouvindo)
#lb_ouvindo.pack()


#Criar Botão
bt_falar = Button(fr_inicial, bd=0, image=img_botao, command=lambda:conversa(tela_inicial,fr_conversa))
bt_falar.place(width=220, height=92, x=140, y=447)


#Iniciar a tela
fr_inicial.tkraise()
#fr_conversa.tkraise()
bt_falar.tkraise()
tela_inicial.mainloop()

#==========================================================================================================================================

# CRIAR UMA CLASSE CHAMADA tarefas() ONDE VAI CONTER TODAS AS FUNÇÕES QUE IRAM EXECUTAR PROCESSOS COMPLEXOS

#A depender da função, se precisar sair da página, já finalizar direto com um comando tipo end do vba: (Finalizar código)!