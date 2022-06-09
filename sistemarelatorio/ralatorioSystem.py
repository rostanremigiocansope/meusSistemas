from tkinter import*
from tkinter  import ttk
import tkinter.messagebox
#=========atac=============
from tkinter import filedialog
attachments = []
#====Enviar Imail======================>
import smtplib
from email.message import EmailMessage
#======================================>
from PIL import ImageTk, Image


def salvar():       
        a=texto.get("1.0", "end")
        arquivo = open("Informacoes_Adicionais.txt", 'a', encoding="utf-8")
        arquivo.write(f"{a}")
        arquivo.write("\n")
        arquivo.close()
        texto.delete('1.0', "end")
        textoVisualisar.delete('1.0', "end")
        tkinter.messagebox.showinfo("INFORMÇÕES","Informação Salva!!!")

def visualisar():
        try:
                arquivo = open("Informacoes_Adicionais.txt", encoding="utf-8")
                conteudo = arquivo.read()
                textoVisualisar.insert(END,conteudo)
        except:
            tkinter.messagebox.showinfo("ERRO","Não tem informações Adicionais!!!")

def limpar():
        textoVisualisar.delete('1.0', "end")

def excluir1():
        res=tkinter.messagebox.askyesno("Excluir","Confirma excluir as Informações adicionais!")
        if(res==True):
                arquivo = open("Informacoes_Adicionais.txt", 'w', encoding="utf-8")
                arquivo.write("")
                arquivo.close()
                limpar()

def DocWord():
        if NumrelatEd.get()=="":
                tkinter.messagebox.showinfo(title="ERRO", message="Prencha todos Campos")
        else:
                try:
                        
                        relat =relatorio.get()
                        Ult=ultimas.get()
                        Comp=compreendida.get()
                        Dodia=doDia.get()
                        Domes=doMes.get()
                        Doano=doAno.get()

                        HorasSeg=horasSeguinte.get()
                        DataSeg=dataSeguinte.get()
                        MesSeg=mesSeguinte.get()
                        AnoSeg=anoSeguinte.get()

                        SitInterna=sitInterna.get()
                        SitTropa=sitTropa.get()
                        SaudeMilitar=saudeMilitar.get()
                        DiscMilitar=discMilitar.get()
                        Presos=presos.get()

                        relaInfAdicionais=texto3.get("1.0", "end")

                        SuperCess=superCessante.get()
                        OpCess=opCessante.get()
                        AdjCess=adjuntoCessante.get()
                        GuardCess=guardaCessante.get()
                        ObsCess=textoObs1.get("1.0", "end")

                        SuperSuce=superSucessor.get()
                        OpSuce=opSucessor.get()
                        AdjSuce=adjuntoSucessor.get()
                        GuardSuce=guardaSucessor.get()
                        ObsSuce=textoObs2.get("1.0", "end")
                        
                        arquivo = open(f"{relat}.docx", 'a', encoding="utf-8")
                        arquivo.write("Republica de Moçambique")
                        arquivo.write("\n")
                        arquivo.write("Ministério da defesa Nacional")
                        arquivo.write("\n")
                        arquivo.write("Estado Maior General")
                        arquivo.write("\n")
                        arquivo.write("Comando do Exercito")
                        arquivo.write("\n")
                        arquivo.write("Brigada de Infantaria de Cuamba")
                        arquivo.write("\n")
                        arquivo.write("Posto de Oficial dia Operacional")
                        arquivo.write("\n\n")

                        relaData=relatorioData.get()
                        relaMes=relatorioMes.get()
                        relaAno=relatorioAno.get()
                        arquivo.write(f"Cuamba aos {relaData} de {relaMes} de {relaAno}")
                        arquivo.write("\n\n")

                        
                        
                        arquivo.write(f"O presente informe tem como objectivo de reportar as actividades realizadas ")
                        arquivo.write(f"nas ultimas {Ult} horas no periodo compreendido entre {Comp} horas do dia {Dodia}/{Domes}/{Doano}")
                        arquivo.write(f" e as {HorasSeg} do dia {DataSeg}/{MesSeg}/{AnoSeg}, em cumprimento do sistema ")
                        arquivo.write(f"informativo em vigor nas FADM.")
                        arquivo.write("\n\n")

                        arquivo.write(f"1. A situação interna {SitInterna}.")
                        arquivo.write("\n")
                        arquivo.write(f"2. A situação das nossas tropas ao longo do periodo foi tida como {SitTropa}.")
                        arquivo.write("\n")
                        arquivo.write(f"3. Saúde militar {SaudeMilitar}.")
                        arquivo.write("\n")
                        arquivo.write(f"4. Disciplina Militar caracterizou-se {DiscMilitar}.")
                        arquivo.write("\n")
                        arquivo.write(f"5. Presos/Detidos {Presos}.")
                        arquivo.write("\n\n\n")

                        arquivo.write("Informações Adicionais")
                        arquivo.write("\n\n")
                        arquivo.write(f"{relaInfAdicionais}")
                        arquivo.write("\n\n\n")

                        arquivo.write("Equipe de Serviço Cessante")
                        arquivo.write("\n\n")
                        arquivo.write(f"1. Oficial dia supervisor: {SuperCess};")
                        arquivo.write("\n")
                        arquivo.write(f"2. Oficial dia Operacional: {OpCess};")
                        arquivo.write("\n")
                        arquivo.write(f"3. Oficial dia Adjunto: {AdjCess};")
                        arquivo.write("\n")
                        arquivo.write(f"4. Comandante da Guarda: {GuardCess};")
                        arquivo.write("\n")
                        arquivo.write(f"{ObsCess}")
                        arquivo.write("\n\n\n")


                        arquivo.write("Equipe de Serviço Sucessor")
                        arquivo.write("\n\n")
                        arquivo.write(f"1. Oficial dia supervisor: {SuperSuce};")
                        arquivo.write("\n")
                        arquivo.write(f"2. Oficial dia Operacional: {OpSuce};")
                        arquivo.write("\n")
                        arquivo.write(f"3. Oficial dia Adjunto: {AdjSuce};")
                        arquivo.write("\n")
                        arquivo.write(f"4. Comandante da Guarda: {GuardSuce};")
                        arquivo.write("\n")
                        arquivo.write(f"{ObsSuce}")

                        arquivo.close()
                        tkinter.messagebox.showinfo("Relatório","Relatório Criado Com Sucesso!!!")
                        limpar1()
                except:
                        tkinter.messagebox.showinfo("ERRO","Feixar o documento Word Aberto!!!")

def attachFile():
        filename = filedialog.askopenfilename(initialdir="c:/", title="please select a file")
        attachments.append(filename)

def enviar():
       try:
                msg      = EmailMessage()
                username = imail.get()
                password = senha.get()
                to       = destino.get()
                subject  = assunto.get()
                body     = mensagem.get()
                msg["subject"] = subject
                msg["from"] = username
                msg["to"] = to
                msg.set_content(body)

                for filename in attachments:
                    filetype = filename.split(".")
                    filetype = filetype[1]
                    if filetype == "jpg" or filetype =="JPG" or filetype == "png" or filetype =="PNG":
                        import imghdr
                        with open(filename, "rb") as f:
                            file_data = f.read()
                            image_type = imghdr.what(filename)
                        msg.add_attachment(file_data, maintype="image", sutype=image_type, filename=f.name)
                    else:
                        with open(filename, "rb") as f:
                            file_data = f.read()
                        msg.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=f.name)
                
                if username=="" or password=="" or to=="" or subject=="" or body=="":
                    notif.config(text="Todos Campos Obrigatórios", fg="red")
                    return
                else:
                    server = smtplib.SMTP("smtp.gmail.com", 587)
                    server.starttls()
                    server.login(username, password)
                    server.send_message(msg)
                    notif.config(text="Email enviado", fg="blue")
       except:
                notif.config(text="Erro ao enviar email", fg="red")

def limpar1():
        NumrelatEd.delete(0, "end")
        localEd.delete(0, "end")
        Data.set("")
        Mes.set("")
        Ano.set("")
        edinfo1horas1.delete(0, "end")
        edinfo2horas.delete(0, "end")
        Datarelatorio1.set("")
        Mesrelatorio1.set("")
        Anorelatorio1.set("")
        edinfo1horas2.delete(0, "end")
        Datarelatorio2.set("")
        Mesrelatorio2.set("")
        Anorelatorio2.set("")
        Edinfo7.delete(0, "end")
        Edinfo8.delete(0, "end")
        Edinfo9.delete(0, "end")
        Edinfo10.delete(0, "end")
        Edinfo11.delete(0, "end")
        texto3.delete('1.0', "end")
        EdSupervisor1.delete(0, "end")
        EdOperacional1.delete(0, "end")
        EdAdjunto1.delete(0, "end")
        EdGuarda1.delete(0, "end")
        textoObs1.delete('1.0', "end")
        EdSupervisor2.delete(0, "end")
        EdOperacional2.delete(0, "end")
        EdAdjunto2.delete(0, "end")
        EdGuarda2.delete(0, "end")
        textoObs2.delete('1.0', "end")

def remover():
        EdImail.delete(0, "end")
        EdSenha.delete(0, "end")
        EdDestino.delete(0, "end")
        EdAssunto.delete(0, "end")
        EdMensagem.delete(0, "end")

def sair():
        res=tkinter.messagebox.askyesno("Sair","Deseja sair do Programa!")
        if(res==True):
                janela.destroy()

def anexo():
            filename = filedialog.askopenfilename(initialdir="c:/", title="please select a file")
            attachments.append(filename)
            notif.config(fg="blue", text ="Adicionado" + str(len(attachments)) + " Arquivos")



janela = Tk()

janela.geometry("1350x700+0+0")

janela.title("Sistema Cansope")
janela.iconbitmap('C:/Users/Rostan/Desktop/Cuamba/SistemaRelatorio/ralatorioSystem/imagemC/icon.ico')

#===============================FPrincipal============================================
framePrincipal = Frame(janela, width=1350, height=700, bd=5, relief="raise")
framePrincipal.grid()
#===============================FrameTitulo==================================
lblFrameTitulo = LabelFrame(framePrincipal,bg="teal", width=1320, height=50)
lblFrameTitulo.place(x=10, y=0)

infoTitulo =Label(lblFrameTitulo,font=("Times new roman", 20,"bold"),bg="teal",fg="white", text="Informe Diário do Posto de Oficial dia Operacional/Niassa/Cuamba")
infoTitulo.place(x=100,y=0)

infoautor1 =Label(lblFrameTitulo,font=("Times new roman", 10),bg="teal",fg="white", text="Desenvolvedor do Sistema:")
infoautor1.place(x=1000,y=0)
infoautor2 =Label(lblFrameTitulo,font=("Times new roman", 10),bg="teal",fg="white", text="Rostan Peter Remigio Cansope(Tenente Art. Terrestre)")
infoautor2.place(x=1000,y=20)
#===============================FrameInformaçoesAdicionais==================
lblFrameInfAd = LabelFrame(framePrincipal,bg="teal",fg="white", width=550, height=600)
lblFrameInfAd.place(x=10, y=60)

infoAd =Label(lblFrameInfAd,font=("Times new roman", 15,"bold","underline"),bg="teal",fg="white", text="INFORMAÇÕES ADICIONAIS")
infoAd.place(x=10,y=5)
texto = Text(lblFrameInfAd,font=("Times new roman", 15),bg="white", width=50, height=4,  bd=5, relief=RIDGE)
texto.place(x=10, y=40)
btnSalvar = Button(lblFrameInfAd, text="Salvar",width=10,bg="purple",fg="white", bd=3, relief="raise",command=salvar)
btnSalvar.place(x=440, y=150)


btnVisualisar = Button(lblFrameInfAd, text="Visualisar",bg="purple",fg="white", width=10, bd=3, relief="raise",command=visualisar)
btnVisualisar.place(x=10, y=200)

btnLimpar = Button(lblFrameInfAd, text="Limpar", width=10,bg="purple",fg="white", bd=3, relief="raise",command=limpar)
btnLimpar.place(x=100, y=200)

btnExcluir = Button(lblFrameInfAd, text="Excluir",width=10,bg="purple",fg="white", bd=3, relief="raise",command=excluir1)
btnExcluir.place(x=190, y=200)

textoVisualisar = Text(lblFrameInfAd,font=("Times new roman",11),bg="white", width=71, height=19,  bd=5, relief=RIDGE)
textoVisualisar.place(x=10, y=240)
#===============================RelatorioFrame============================================

relatorio=StringVar()
local=StringVar()
relatorioData=StringVar()
relatorioMes=StringVar()
relatorioAno=StringVar()

lblFrameRelat = LabelFrame(framePrincipal,bg="teal",fg="white", width=760, height=600)
lblFrameRelat.place(x=570, y=60)

NumrelatLabel = Label(lblFrameRelat,font=("Times new roman", 15,"bold","underline"),bg="teal",fg="white", text="RELATÓRIO No")
NumrelatLabel.place(x=10, y=10)
NumrelatEd=Entry(lblFrameRelat,font=("Times new roman", 15), width=7, textvariable=relatorio)
NumrelatEd.place(x=170, y=10)

localEd=Entry(lblFrameRelat,font=("Times new roman", 15), width=10, textvariable=local)
localEd.place(x=360, y=10)
diaData=Label(lblFrameRelat,font=("Times new roman", 12),bg="teal",fg="white", text="aos")
diaData.place(x=470, y=10)

Data=ttk.Combobox(lblFrameRelat, state="readonly",
                                           font=("Times New Roman",12,"bold"), width=4, textvariable=relatorioData)
Data["value"]=("","01","02","03","04","05","06","07","08","09","10","11","12",
                             "13","14","15","16","18","19","20","21","22","23","24","25","26",
                             "27","28","29","30","31")
Data.current(0)
Data.place(x=500, y=10)

Mes=ttk.Combobox(lblFrameRelat, state="readonly",
                                           font=("Times New Roman",12,"bold"), width=7, textvariable=relatorioMes)
Mes["value"]=("","Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro")
Mes.current(0)
Mes.place(x=560, y=10)

Ano=ttk.Combobox(lblFrameRelat, state="readonly",
                                           font=("Times New Roman",12,"bold"), width=6, textvariable=relatorioAno)
Ano["value"]=("","2022","2023","2024","2025","2026","2027","2028","2029","2030","2031","2032","2033","2034","2035","2036"
              ,"2037","2038","2039","2040","2041","2042","2043","2044","2045","2046","2047","2048","2049","2050")

Ano.current(0)
Ano.place(x=645, y=10)

ultimas=StringVar()
compreendida=StringVar()
doDia=StringVar()
doMes=StringVar()
doAno=StringVar()


lblinfo1=Label(lblFrameRelat,font=("Times New Roman",12,"bold"),bg="teal",fg="white",text="Informe das últimas")
lblinfo1.place(x=10,y=40)
edinfo1horas1=Entry(lblFrameRelat,font=("Times New Roman",12,"bold"),width=3, textvariable=ultimas)
edinfo1horas1.place(x=155,y=40)
lblinfo2=Label(lblFrameRelat,font=("Times New Roman",12,"bold"),bg="teal",fg="white",text="horas, no período compreendido entre")
lblinfo2.place(x=180,y=40)
edinfo2horas=Entry(lblFrameRelat,font=("Times New Roman",12,"bold"),width=5, textvariable=compreendida)
edinfo2horas.place(x=440,y=40)
lblinfo3=Label(lblFrameRelat,font=("Times New Roman",12,"bold"),bg="teal",fg="white",text=" do dia")
lblinfo3.place(x=490,y=40)


Datarelatorio1=ttk.Combobox(lblFrameRelat, state="readonly",font=("Times New Roman",12,"bold"), width=4, textvariable=doDia)
Datarelatorio1["value"]=("","01","02","03","04","05","06","07","08","09","10","11","12","13","14","15","16","18","19","20","21","22","23","24","25","26",
                             "27","28","29","30","31")
Datarelatorio1.current(0)
Datarelatorio1.place(x=560, y=40)

Mesrelatorio1=ttk.Combobox(lblFrameRelat, state="readonly",font=("Times New Roman",12,"bold"), width=4, textvariable=doMes)
Mesrelatorio1["value"]=("","01","02","03","04","05","06","07","08","09","10","11","12")
Mesrelatorio1.current(0)
Mesrelatorio1.place(x=620, y=40)

Anorelatorio1=ttk.Combobox(lblFrameRelat, state="readonly",font=("Times New Roman",12,"bold"), width=6, textvariable=doAno)
Anorelatorio1["value"]=("","2022","2023","2024","2025","2026","2027","2028","2029","2030","2031","2032","2033","2034","2035","2036"
              ,"2037","2038","2039","2040","2041","2042","2043","2044","2045","2046","2047","2048","2049","2050")
Anorelatorio1.current(0)
Anorelatorio1.place(x=680, y=40)

horasSeguinte=StringVar()
dataSeguinte=StringVar()
mesSeguinte=StringVar()
anoSeguinte=StringVar()

lblinfo4=Label(lblFrameRelat,font=("Times New Roman",12,"bold"),bg="teal",fg="white",text="e as")
lblinfo4.place(x=10,y=70)
edinfo1horas2=Entry(lblFrameRelat,font=("Times New Roman",12,"bold"),width=5,textvariable=horasSeguinte)
edinfo1horas2.place(x=50,y=70)
lblinfo5=Label(lblFrameRelat,font=("Times New Roman",12,"bold"),bg="teal",fg="white",text=" do dia")
lblinfo5.place(x=100,y=70)

Datarelatorio2=ttk.Combobox(lblFrameRelat, state="readonly",font=("Times New Roman",12,"bold"), width=4, textvariable=dataSeguinte)
Datarelatorio2["value"]=("","01","02","03","04","05","06","07","08","09","10","11","12","13","14","15","16","18","19","20","21","22","23","24","25","26",
                             "27","28","29","30","31")
Datarelatorio2.current(0)
Datarelatorio2.place(x=170, y=70)

Mesrelatorio2=ttk.Combobox(lblFrameRelat, state="readonly",font=("Times New Roman",12,"bold"), width=4, textvariable=mesSeguinte)
Mesrelatorio2["value"]=("","01","02","03","04","05","06","07","08","09","10","11","12")
Mesrelatorio2.current(0)
Mesrelatorio2.place(x=230, y=70)

Anorelatorio2=ttk.Combobox(lblFrameRelat, state="readonly",font=("Times New Roman",12,"bold"), width=6, textvariable=anoSeguinte)
Anorelatorio2["value"]=("","2022","2023","2024","2025","2026","2027","2028","2029","2030","2031","2032","2033","2034","2035","2036"
              ,"2037","2038","2039","2040","2041","2042","2043","2044","2045","2046","2047","2048","2049","2050")
Anorelatorio2.current(0)
Anorelatorio2.place(x=290, y=70)

lblinfo6=Label(lblFrameRelat,font=("Times New Roman",12,"bold"),bg="teal",fg="white",text="em cumprimento do sistema informativo das FADM.")
lblinfo6.place(x=370,y=70)
#=====
sitInterna=StringVar()
sitTropa=StringVar()
saudeMilitar=StringVar()
discMilitar=StringVar()
presos=StringVar()

lblinfo7=Label(lblFrameRelat,font=("Times New Roman",12,"bold"),bg="teal",fg="white",text="A Situação interna")
lblinfo7.place(x=10,y=110)
Edinfo7=Entry(lblFrameRelat,font=("Times New Roman",12,"bold"),width=21, textvariable=sitInterna)
Edinfo7.place(x=140,y=110)

lblinfo8=Label(lblFrameRelat,font=("Times New Roman",12,"bold"),bg="teal",fg="white",text="; Situação das nossas tropas ao longo do período foi tida como")
lblinfo8.place(x=310,y=110)
Edinfo8=Entry(lblFrameRelat,font=("Times New Roman",12,"bold"),width=21, textvariable=sitTropa)
Edinfo8.place(x=10,y=140)

lblinfo9=Label(lblFrameRelat,font=("Times New Roman",12,"bold"),bg="teal",fg="white",text="; Saúde militar")
lblinfo9.place(x=150,y=140)
Edinfo9=Entry(lblFrameRelat,font=("Times New Roman",12,"bold"),width=58, textvariable=saudeMilitar)
Edinfo9.place(x=260,y=140)

lblinfo10=Label(lblFrameRelat,font=("Times New Roman",12,"bold"),bg="teal",fg="white",text="Disciplina Militar caracterizou-se")
lblinfo10.place(x=10,y=170)
Edinfo10=Entry(lblFrameRelat,font=("Times New Roman",12,"bold"),width=21, textvariable=discMilitar)
Edinfo10.place(x=240,y=170)

lblinfo11=Label(lblFrameRelat,font=("Times New Roman",12,"bold"),bg="teal",fg="white",text="; Presos")
lblinfo11.place(x=410,y=170)
Edinfo11=Entry(lblFrameRelat,font=("Times New Roman",12,"bold"),width=31, textvariable=presos)
Edinfo11.place(x=475,y=170)
#=========
lblinfo12=Label(lblFrameRelat,font=("Times New Roman",12,"bold","underline"),bg="teal",fg="white",text="INFORMAÇÕES ADICIONAIS")
lblinfo12.place(x=10,y=210)
texto3 = Text(lblFrameRelat,font=("Times new roman", 10),bg="white", width=120, height=6,  bd=5, relief=RIDGE)
texto3.place(x=10, y=240)
#========
lblFrameCessante = LabelFrame(lblFrameRelat,bg="teal", width=240, height=230)
lblFrameCessante.place(x=10, y=350)
lblCessante = Label(lblFrameCessante,font=("Times new roman", 14,"bold","underline"),bg="teal",fg="white", text="Equipe de Serviço Cessante")
lblCessante.place(x=0, y=0)

superCessante=StringVar()
opCessante=StringVar()
adjuntoCessante=StringVar()
guardaCessante=StringVar()

superSucessor=StringVar()
opSucessor=StringVar()
adjuntoSucessor=StringVar()
guardaSucessor=StringVar()


lblSupervisor1=Label(lblFrameCessante,font=("Times New Roman",12,"bold"),bg="teal",fg="white",text="1.Of/D/Sup")
lblSupervisor1.place(x=0,y=25)
EdSupervisor1=Entry(lblFrameCessante,font=("Times New Roman",10,"bold"),width=21, textvariable=superCessante)
EdSupervisor1.place(x=80,y=28)

lblOperacional1=Label(lblFrameCessante,font=("Times New Roman",12,"bold"),bg="teal",fg="white",text="2.Of/D/Op")
lblOperacional1.place(x=0,y=50)
EdOperacional1=Entry(lblFrameCessante,font=("Times New Roman",10,"bold"),width=21, textvariable=opCessante)
EdOperacional1.place(x=80,y=53)

lblAdjunto1=Label(lblFrameCessante,font=("Times New Roman",12,"bold"),bg="teal",fg="white",text="3.Of/D/Adj")
lblAdjunto1.place(x=0,y=75)
EdAdjunto1=Entry(lblFrameCessante,font=("Times New Roman",10,"bold"),width=21, textvariable=adjuntoCessante)
EdAdjunto1.place(x=80,y=78)

lblGuarda1=Label(lblFrameCessante,font=("Times New Roman",12,"bold"),bg="teal",fg="white",text="4.C/Guarda")
lblGuarda1.place(x=0,y=100)
EdGuarda1=Entry(lblFrameCessante,font=("Times New Roman",10,"bold"),width=21, textvariable=guardaCessante)
EdGuarda1.place(x=80,y=103)

lblObs=Label(lblFrameCessante,font=("Times New Roman",12,"bold"),bg="teal",fg="white",text="Obs.")
lblObs.place(x=0,y=125)

textoObs1 = Text(lblFrameCessante,font=("Times new roman", 10),bg="white", width=37, height=4,  bd=5, relief=RIDGE)
textoObs1.place(x=0, y=150)
#==
lblFrameSucessor = LabelFrame(lblFrameRelat,bg="teal", width=240, height=230)
lblFrameSucessor.place(x=250, y=350)

lblSucessor = Label(lblFrameSucessor,font=("Times new roman", 14,"bold","underline"),bg="teal",fg="white", text="Equipe de Serviço Sucessor")
lblSucessor.place(x=0, y=0)


lblSupervisor2=Label(lblFrameSucessor,font=("Times New Roman",12,"bold"),bg="teal",fg="white",text="1.Of/D/Sup")
lblSupervisor2.place(x=0,y=25)
EdSupervisor2=Entry(lblFrameSucessor,font=("Times New Roman",10,"bold"),width=21, textvariable=superSucessor)
EdSupervisor2.place(x=80,y=28)

lblOperacional2=Label(lblFrameSucessor,font=("Times New Roman",12,"bold"),bg="teal",fg="white",text="2.Of/D/Op")
lblOperacional2.place(x=0,y=50)
EdOperacional2=Entry(lblFrameSucessor,font=("Times New Roman",10,"bold"),width=21, textvariable=opSucessor)
EdOperacional2.place(x=80,y=53)

lblAdjunto2=Label(lblFrameSucessor,font=("Times New Roman",12,"bold"),bg="teal",fg="white",text="3.Of/D/Adj")
lblAdjunto2.place(x=0,y=75)
EdAdjunto2=Entry(lblFrameSucessor,font=("Times New Roman",10,"bold"),width=21, textvariable=adjuntoSucessor)
EdAdjunto2.place(x=80,y=78)

lblGuarda2=Label(lblFrameSucessor,font=("Times New Roman",12,"bold"),bg="teal",fg="white",text="4.C/Guarda")
lblGuarda2.place(x=0,y=100)
EdGuarda2=Entry(lblFrameSucessor,font=("Times New Roman",10,"bold"),width=21, textvariable=guardaSucessor)
EdGuarda2.place(x=80,y=103)

lblObs=Label(lblFrameSucessor,font=("Times New Roman",12,"bold"),bg="teal",fg="white",text="Obs.")
lblObs.place(x=0,y=125)

textoObs2 = Text(lblFrameSucessor,font=("Times new roman", 10),bg="white", width=37, height=4,  bd=5, relief=RIDGE)
textoObs2.place(x=0, y=150)
#====
btnDocWord = Button(lblFrameRelat, text="DocWord",bg="purple",fg="white", width=7, bd=3, relief="raise",command=DocWord)
btnDocWord.place(x=500, y=350)

btnVisualisar = Button(lblFrameRelat, text="Visualisar",bg="purple",fg="white", width=7, bd=3, relief="raise",command=attachFile)
btnVisualisar.place(x=560, y=350)

btnLimpar = Button(lblFrameRelat, text="Limpar",bg="purple",fg="white", width=7, bd=3, relief="raise",command=limpar1)
btnLimpar.place(x=620, y=350)

btnSair = Button(lblFrameRelat, text="Sair",bg="red",fg="white", width=7, bd=3, relief="raise",command=sair)
btnSair.place(x=680, y=350)
#===

lblFrameImail = LabelFrame(lblFrameRelat,bg="teal", width=240, height=200)
lblFrameImail.place(x=500, y=380)

imail=StringVar()
senha=StringVar()
destino=StringVar()
assunto=StringVar()
mensagem=StringVar()


lblImail=Label(lblFrameImail,font=("Times New Roman",12,"bold"),bg="teal",fg="white",text="E-mail")
lblImail.place(x=0,y=0)
EdImail=Entry(lblFrameImail,font=("Times New Roman",10,"bold"),width=21,textvariable=imail)
EdImail.place(x=80,y=3)

lblSenha=Label(lblFrameImail,font=("Times New Roman",12,"bold"),bg="teal",fg="white",text="Senha")
lblSenha.place(x=0,y=25)
EdSenha=Entry(lblFrameImail,font=("Times New Roman",10,"bold"),width=21, textvariable=senha)
EdSenha.place(x=80,y=28)

lblDestino=Label(lblFrameImail,font=("Times New Roman",12,"bold"),bg="teal",fg="white",text="Destino")
lblDestino.place(x=0,y=50)
EdDestino=Entry(lblFrameImail,font=("Times New Roman",10,"bold"),width=21, textvariable=destino)
EdDestino.place(x=80,y=53)

lblAssunto=Label(lblFrameImail,font=("Times New Roman",12,"bold"),bg="teal",fg="white",text="Assunto")
lblAssunto.place(x=0,y=75)
EdAssunto=Entry(lblFrameImail,font=("Times New Roman",10,"bold"),width=21, textvariable=assunto)
EdAssunto.place(x=80,y=78)

lblMensagem=Label(lblFrameImail,font=("Times New Roman",12,"bold"),bg="teal",fg="white",text="Mensagem")
lblMensagem.place(x=0,y=100)
EdMensagem=Entry(lblFrameImail,font=("Times New Roman",10,"bold"),width=21, textvariable=mensagem)
EdMensagem.place(x=80,y=103)

notif =Label(lblFrameImail,width=20, text="",bg="gray", font=("Times New Roman", 10))
notif.place(x=30, y=130)
#========
btnAnexo = Button(lblFrameImail, text="Anexo",bg="purple",fg="white", width=9, bd=3, relief="raise",command=anexo)
btnAnexo.place(x=0, y=160)

btnEnviar = Button(lblFrameImail, text="Enviar",bg="purple",fg="white", width=9, bd=3, relief="raise",command=enviar)
btnEnviar.place(x=80, y=160)

btnLimpar = Button(lblFrameImail, text="Limpar",bg="purple",fg="white", width=9, bd=3, relief="raise",command=remover)
btnLimpar.place(x=160, y=160)


janela.mainloop()
