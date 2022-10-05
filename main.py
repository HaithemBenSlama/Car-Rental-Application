from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import sys
from PyQt5.uic import loadUiType
import re
from PyQt5 import *
from PyQt5 import QtCore
from PyQt5 import QtGui
from Classes import *
from MessageBoxes import *
import random
from openpyxl import *
import webbrowser

ui, _ = loadUiType('MainUi.ui')
ui2, _ = loadUiType('Loading.ui')
counter=0
class LoadingScreen(QMainWindow,ui2):
    def __init__(self):
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.setWindowFlag(Qt.FramelessWindowHint)
        self.show()
        ## QTIMER ==> START
        self.timer = QtCore.QTimer()
        self.timer.timeout.connect(self.progress)
        # TIMER IN MILLISECONDS
        self.timer.start(35)
    def progress(self):
        global counter
        # SET VALUE TO PROGRESS BAR
        self.progressBar.setValue(counter)

        # CLOSE SPLASH SCREE AND OPEN APP
        if counter > 100:
            # STOP TIMER
            self.timer.stop()
            # SHOW MAIN WINDOW
            self.main = MainApp()
            self.main.show()
            # CLOSE SPLASH SCREEN
            self.close()
        # INCREASE COUNTER
        counter += 1
        
class MainApp(QMainWindow,ui):
    Data=Main()
    recupClient=False
    recupVoiture=False
    recupLocation=False
    saveClient=True
    saveVoiture=True
    saveLocation=True
    
    #function btn_client
    def click_btn_client(self):
        if(self.recupClient==False):
            msgBox=MessageBoxErreur()
            msgBox.setText("Erreur: \n\nEssayer de récupérer le fichier Client.csv\n")
            msgBox.exec()
        else:
            self.click_btn_add_client()
            self.stacked_dash.setCurrentIndex(0)
        
    #function btn_voiture
    def click_btn_voiture(self):
        if(self.recupVoiture==False):
            msgBox=MessageBoxErreur()
            msgBox.setText("Erreur: \n\nEssayer de récupérer le fichier Voiture.csv\n")
            msgBox.exec()
        else:
            self.click_btn_add_voiture()
            self.stacked_dash.setCurrentIndex(1)
        
    #function btn_location
    def click_btn_location(self):
        if(self.recupLocation==False):
            msgBox=MessageBoxErreur()
            msgBox.setText("Erreur: \n\nEssayer de récupérer le fichier Location.csv\n")
            msgBox.exec()
        else:
            self.click_btn_add_location()
            self.stacked_dash.setCurrentIndex(3)
    #function btn_load_client
    def click_btn_load_client(self):
        if(self.recupClient==True):
            msg=MessageBoxErreur()
            msg.setText("Erreur: \n\nVous Avez Deja Récuperé le fichier Client.csv\n")
            msg.exec()
        else:
            self.Data.load_client()
            self.recupClient=True
            msg=MessageBoxValid()
            msg.setText("Information: \n\nLe Fichier Client.csv est récupéré avec succès\n")
            msg.exec()
    #function btn_load_car
    def click_btn_load_car(self):
        if(self.recupVoiture==True):
            msg=MessageBoxErreur()
            msg.setText("Erreur: \n\nVous Avez Deja Récuperé le fichier Voiture.csv\n")
            msg.exec()
        else:
            self.Data.load_car()
            self.recupVoiture=True
            msg=MessageBoxValid()
            msg.setText("Information: \n\nLe Fichier Voiture.csv est récupéré avec succès\n")
            msg.exec()

    #function btn_load_location
    def click_btn_load_location(self):
        if(self.recupLocation==True):
            msg=MessageBoxErreur()
            msg.setText("Erreur: \n\nVous Avez Deja Récuperé le fichier Location.csv\n")
            msg.exec()
        else:
            self.Data.load_Location()
            self.recupLocation=True
            msg=MessageBoxValid()
            msg.setText("Information: \n\nLe Fichier Location.csv est récupéré avec succès\n")
            msg.exec()

    #function btn_load_all
    def click_btn_load_all(self):
        self.click_btn_load_car()
        self.click_btn_load_client()
        self.click_btn_load_location()
        self.click_btn_accueil()

    #function btn_save_client
    def click_btn_save_client(self):
        if(self.recupClient==False):
            msg=MessageBoxErreur()
            msg.setText("Erreur: \n\nEssayer de récupérer le fichier Client.csv D'abord\n")
            msg.exec()
        else:
            self.Data.save_client()
            self.saveClient=True
            msg=MessageBoxValid()
            msg.setText("Information: \n\nLe Fichier Client.csv est sauvgardé avec succès\n")
            msg.exec()

    #function btn_save_car
    def click_btn_save_car(self):
        if(self.recupVoiture==False):
            msg=MessageBoxErreur()
            msg.setText("Erreur: \n\nEssayer de récupérer le fichier Voiture.csv D'abord\n")
            msg.exec()
        else:
            self.Data.save_car()
            self.saveVoiture=True
            msg=MessageBoxValid()
            msg.setText("Information: \n\nLe Fichier Voiture.csv est sauvgardé avec succès\n")
            msg.exec()

    #function btn_save_location
    def click_btn_save_location(self):
        if(self.recupLocation==False):
            msg=MessageBoxErreur()
            msg.setText("Erreur: \n\nEssayer de récupérer le fichier Location.csv D'abord\n")
            msg.exec()
        else:
            self.Data.save_Location()
            self.saveLocation=True
            msg=MessageBoxValid()
            msg.setText("Information: \n\nLe Fichier Location.csv est sauvgardé avec succès\n")
            msg.exec()


    #function btn_save_all
    def click_btn_save_all(self):
        self.click_btn_save_car()
        self.click_btn_save_client()
        self.click_btn_save_location()

    #function btn_add_location_next
    def click_btn_add_location_next(self):
        msgBox=MessageBoxErreur()
        dateNow=QDateTime.currentDateTime()
        cin=self.line_cin_add_loc.text()
        datec=self.dateTime_add_loc.dateTime()
        duree=self.line_dure_add_loc.text()
        s1,s2,s3,s4=True,True,True,True
        styleNormal="border-radius:10px;border: 2px solid rgb(133, 183, 198);background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        styleErreur="border-radius:10px;border: 2px solid red;background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        if(len(cin)==8 and cin not in self.Data.dictClient ):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ CIN\t\n*Pas de Client avec Cette CIN\n\n==> Ajouter un Client avec Cette CIN\n==> Vérifier le CIN")
            msgBox.exec()
            self.line_cin_add_loc.setClearButtonEnabled(True)
            self.line_cin_add_loc.setStyleSheet(styleErreur)
            self.btn_add_client_location.setVisible(True)
            s2=False
        else:
            self.line_cin_add_loc.setClearButtonEnabled(False)
            self.line_cin_add_loc.setStyleSheet(styleNormal)
            self.btn_add_client_location.setVisible(False)
            s2=True
        if(len(cin)<8):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ CIN\t\n*Champs Obligatoire\n*8 Chiffres\n")
            msgBox.exec()
            self.line_cin_add_loc.setClearButtonEnabled(True)
            self.line_cin_add_loc.setStyleSheet(styleErreur)
            s1=False
        else:
            self.line_cin_add_loc.setClearButtonEnabled(False)
            self.line_cin_add_loc.setStyleSheet(styleNormal)
            s1=True
        if(datec<dateNow):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Date\t\n*Champ Obligatoire\n*Date Doit être supérieur à Date aujourd'hui "+dateNow.toString("dd/MM/yyyy HH'h'"))
            msgBox.exec()
            self.dateTime_add_loc.setStyleSheet(styleErreur)
            s4=False
        else:
            self.dateTime_add_loc.setStyleSheet(styleNormal)
            s4=True
        if(duree=="" or int(duree)<1):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Durée\t\n*Champ Obligatoire\n*Au Moin 1 jour")
            msgBox.exec()
            self.line_dure_add_loc.setClearButtonEnabled(True)
            self.line_dure_add_loc.setStyleSheet(styleErreur)
            s3=False
        else:
            self.line_dure_add_loc.setClearButtonEnabled(False)
            self.line_dure_add_loc.setStyleSheet(styleNormal)
            s3=True
        
        if(s1 and s2 and s3 and s4):
            self.comboBox_add_location.clear()
            n=[]
            for value in self.Data.dictLocation.values():
                    if(value.date<=datec<=value.date.addDays(int(value.duree)) or datec<=value.date<=datec.addDays(int(duree))):
                        n.append(value.matricule)
            for i,j in self.Data.dictVoiture.items():
                if (i not in n):
                    self.comboBox_add_location.addItem(str(i))
                
            if self.comboBox_add_location.count()==0:
                msgBox.setText("Erreur: \n\nDésolé il n'existe pas de voiture Disponible pour cette date, S'il vous plait essayer de changer la Date de Location ")
                msgBox.exec()
            else:
                self.Main_frame.setCurrentIndex(21)
                self.comboBox_add_location.setCurrentIndex(-1)
        
    #function btn_confirm_add_location
    def click_btn_confirm_add_location(self):
        styleNormal="border-radius:10px;border: 2px solid rgb(133, 183, 198);background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        styleErreur="border-radius:10px;border: 2px solid red;background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        msgBox=MessageBoxErreur()
        if(self.line_montant_loc.text()==''):
            msgBox.setText("Erreur: \n\n*Choisir la matricule de la voiture à louée\n")
            msgBox.exec()
            self.line_montant_loc.setStyleSheet(styleErreur)
        else:
            self.line_montant_loc.setStyleSheet(styleNormal)
            num=self.line_num_add_loc.text()
            cin=self.line_cin_add_loc.text()
            datec=self.dateTime_add_loc.dateTime()
            duree=self.line_dure_add_loc.text()
            matricule=self.comboBox_add_location.currentText()
            montant=self.line_montant_loc.text()
            self.saveLocation=False
            self.Data.add_location(num,cin,matricule,datec,duree,montant)
            msg=MessageBoxValid()
            msg.setText("Information: \n\nLa Location Numéro \n"+str(num)+" est ajoutée avec Sucée")
            msg.exec()
            msgBox = QMessageBox()
            msgBox.setIcon(QMessageBox.Question)
            msgBox.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            msgBox.setWindowFlag(Qt.FramelessWindowHint)
            msgBox.setMinimumWidth(500)
            msgBox.setFixedWidth(600)
            msgBox.setWindowOpacity(0.85)
            msgBox.setText("Question: \n\n Voudriez-vous Imprimer la facture ? \n")
            msgBox.setStyleSheet("QPushButton{border-radius:10px;border:2px solid rgb(99, 126, 193);background-color: rgb(85, 96, 124);font: 87 14pt ""Segoe UI"";width:60px;color: rgb(255, 255, 255);}QPushButton:hover{background-color: rgb(114, 153, 146);}QLabel{font: 87 14pt ""Segoe UI"";color: white;}QMessageBox{border-radius:25px;border:7px solid rgb(95,158,160);background-color: rgb(32, 33, 53);}")
            res=msgBox.exec()
            if(res==QMessageBox.Yes):
                self.Data.imprimer_facture(num)
            self.click_btn_add_location()
            
            
    #function btn_serach_date_location
    def click_btn_serach_date_location(self):
        self.Data.search_loc_date(self.dateTime_search_loc.date(),self.View_loc_date)
        if(self.View_loc_date.rowCount()==0):
            msgBox=MessageBoxErreur()
            msgBox.setText("Erreur: \n\n*Pas de voiture avec cette Date\n")
            msgBox.exec()
    #function btn_add_client_confirm__from_loc
    def click_btn_add_client_confirm__from_loc(self):
        s1,s2,s3,s4,s5,s6,s7=True,True,True,True,True,True,True
        msgBox=MessageBoxErreur()
        regexmail=r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
        styleNormal="border-radius:10px;border: 2px solid rgb(133, 183, 198);background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        styleErreur="border-radius:10px;border: 2px solid red;background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        cin=self.line_Cin_add_2.text()
        nom=self.line_nom_add_2.text()
        prenom=self.line_prenom_add_2.text()
        age=self.line_age_add_2.text()
        adresse=self.line_adresse_add_2.text()
        mail=self.line_mail_add_2.text()
        tel=self.line_numero_add_2.text()
        if(cin=="" or len(cin)<8 or cin in self.Data.dictClient):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ CIN\t\n*Champs Obligatoire\n*8 Chiffres\n*Cin Existe Déja")
            msgBox.exec()
            self.line_Cin_add_2.setClearButtonEnabled(True)
            self.line_Cin_add_2.setStyleSheet(styleErreur)
            s1=False
        else:
            self.line_Cin_add_2.setClearButtonEnabled(False)
            self.line_Cin_add_2.setStyleSheet(styleNormal)
            s1=True
        if(nom==""):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Nom\t\n*Champ Obligatoire")
            msgBox.exec()
            self.line_nom_add_2.setClearButtonEnabled(True)
            self.line_nom_add_2.setStyleSheet(styleErreur)
            s2=False
        else:
            self.line_nom_add_2.setClearButtonEnabled(False)
            self.line_nom_add_2.setStyleSheet(styleNormal)
            s2=True
        if(prenom==""):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Prénom\t\n*Champ Obligatoire")
            msgBox.exec()
            self.line_prenom_add_2.setClearButtonEnabled(True)
            self.line_prenom_add_2.setStyleSheet(styleErreur)
            s3=False
        else:
            self.line_prenom_add_2.setClearButtonEnabled(False)
            self.line_prenom_add_2.setStyleSheet(styleNormal)
            s3=True
        if(age=="" or int(age)<18):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Age\t\n*Champ Obligatoire\n*Age > 18 Ans")
            msgBox.exec()
            self.line_age_add_2.setClearButtonEnabled(True)
            self.line_age_add_2.setStyleSheet(styleErreur)
            s4=False
        else:
            self.line_age_add_2.setClearButtonEnabled(False)
            self.line_age_add_2.setStyleSheet(styleNormal)
            s4=True
        if(adresse==""):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Adresse\t\n*Champ Obligatoire\n")
            msgBox.exec()
            self.line_adresse_add_2.setClearButtonEnabled(True)
            self.line_adresse_add_2.setStyleSheet(styleErreur)
            s5=False
        else:
            self.line_adresse_add_2.setClearButtonEnabled(False)
            self.line_adresse_add_2.setStyleSheet(styleNormal)
            s5=True
        if(mail=="" or not re.fullmatch(regexmail, mail)):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Mail\t\n*Champ Obligatoire\n*Sous La forme: User@domaine.com")
            msgBox.exec()
            self.line_mail_add_2.setClearButtonEnabled(True)
            self.line_mail_add_2.setStyleSheet(styleErreur)
            s6=False
        else:
            self.line_mail_add_2.setClearButtonEnabled(False)
            self.line_mail_add_2.setStyleSheet(styleNormal)
            s6=True
        if(tel=="" or len(tel)<8):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Numéro Tél.\t\n*Champ Obligatoire\n*8 Chiffres")
            msgBox.exec()
            self.line_numero_add_2.setClearButtonEnabled(True)
            self.line_numero_add_2.setStyleSheet(styleErreur)
            s7=False
        else:
            self.line_numero_add_2.setClearButtonEnabled(False)
            self.line_numero_add_2.setStyleSheet(styleNormal)
            s7=True
        if(s1 and s2 and s3 and s4 and s5 and s6 and s7):
            self.saveClient=False
            self.Data.add_client(cin,nom,prenom,age,adresse,mail,tel)
            msg=MessageBoxValid()
            msg.setText("Information: \n\nMr/Mrs "+str(nom)+" "+str(prenom)+" est Ajoutée Avec Succées")
            msg.exec()
            self.Main_frame.setCurrentIndex(20)
            self.line_cin_add_loc.setText(cin)
            self.btn_add_client_location.setVisible(False)
    
    #function btn_add_client_location
    def click_btn_add_client_location(self):
        self.Main_frame.setCurrentIndex(34)
        self.line_Cin_add_2.setText(self.line_cin_add_loc.text())
    
    #function btn_delete_loc_confirm
    def click_btn_delete_loc_confirm(self):
        numero='#'+self.line_numero_del_loc.text()
        styleNormal="border-radius:10px;border: 2px solid rgb(133, 183, 198);background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        styleErreur="border-radius:10px;border: 2px solid red;background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        msgBox=MessageBoxErreur()
        if(numero=="" or numero not in self.Data.dictLocation.keys()):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Numéro\t\n*Champs Obligatoire\n*5 Chiffres\n*Pas de Location avec ce Numéro")
            msgBox.exec()
            self.line_numero_del_loc.setClearButtonEnabled(True)
            self.line_numero_del_loc.setStyleSheet(styleErreur)
        else:
            self.line_numero_del_loc.setClearButtonEnabled(False)
            self.line_numero_del_loc.setStyleSheet(styleNormal)
            self.Data.delete_location(numero)
            msg=MessageBoxValid()
            msg.setText("Information: \n\nFacture N: "+str(numero)+" est supprimée avec Succée")
            msg.exec()
            self.Main_frame.setCurrentIndex(22)

    #function btn_search_loc_cin_confirm
    def click_btn_search_loc_cin_confirm(self):
        cin=self.line_cin_search_loc.text()
        l=[]
        for i,j in self.Data.dictLocation.items():
            if(j.cin==cin):
                l.append(cin)
        styleNormal="border-radius:10px;border: 2px solid rgb(133, 183, 198);background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        styleErreur="border-radius:10px;border: 2px solid red;background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        msgBox=MessageBoxErreur()
        if(cin=="" or len(cin)<8 or cin not in l):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ CIN\t\n*Champs Obligatoire\n*8 Chiffres\n*Pas de Location avec cette cin")
            msgBox.exec()
            self.line_cin_search_loc.setClearButtonEnabled(True)
            self.line_cin_search_loc.setStyleSheet(styleErreur)
        else:
            self.line_cin_search_loc.setClearButtonEnabled(False)
            self.line_cin_search_loc.setStyleSheet(styleNormal)
            self.Data.search_loc_cin(cin,self.View_loc_cin)

    #function btn_search_matricule_loc_confirm
    def click_btn_search_matricule_loc_confirm(self):
        matricule=self.line_Matricule_search_loc.text().upper()
        l=[]
        for i,j in self.Data.dictLocation.items():
            if(j.matricule==matricule):
                l.append(matricule)
        styleNormal="border-radius:10px;border: 2px solid rgb(133, 183, 198);background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        styleErreur="border-radius:10px;border: 2px solid red;background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        msgBox=MessageBoxErreur()
        regexmat=r'\b[0-9][0-9][0-9]TN[0-9][0-9][0-9][0-9]\b'
        if(not re.fullmatch(regexmat,matricule) or matricule=="" or matricule not in l):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Matricule\t\n*SOUS LA FORME: 000 TN 0000\n *Pas de Location avec Cette Matricule\n")
            msgBox.exec()
            self.line_Matricule_search_loc.setClearButtonEnabled(True)
            self.line_Matricule_search_loc.setStyleSheet(styleErreur)
        else:
            self.line_Matricule_search_loc.setClearButtonEnabled(False)
            self.line_Matricule_search_loc.setStyleSheet(styleNormal)
            self.Data.search_loc_matricule(matricule,self.View_loc_matricule)


            



    ################################################### Espace Client #############################################################
    #function btn_add_client_confirm
    def click_btn_add_client_confirm(self):
        s1,s2,s3,s4,s5,s6,s7=True,True,True,True,True,True,True
        msgBox=MessageBoxErreur()
        regexmail=r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
        styleNormal="border-radius:10px;border: 2px solid rgb(133, 183, 198);background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        styleErreur="border-radius:10px;border: 2px solid red;background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        cin=self.line_Cin_add.text()
        nom=self.line_nom_add.text()
        prenom=self.line_prenom_add.text()
        age=self.line_age_add.text()
        adresse=self.line_adresse_add.text()
        mail=self.line_mail_add.text()
        tel=self.line_numero_add.text()
        if(cin=="" or len(cin)<8 or cin in self.Data.dictClient):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ CIN\t\n*Champs Obligatoire\n*8 Chiffres\n*Cin Existe Déja")
            msgBox.exec()
            self.line_Cin_add.setClearButtonEnabled(True)
            self.line_Cin_add.setStyleSheet(styleErreur)
            s1=False
        else:
            self.line_Cin_add.setClearButtonEnabled(False)
            self.line_Cin_add.setStyleSheet(styleNormal)
            s1=True
        if(nom==""):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Nom\t\n*Champ Obligatoire")
            msgBox.exec()
            self.line_nom_add.setClearButtonEnabled(True)
            self.line_nom_add.setStyleSheet(styleErreur)
            s2=False
        else:
            self.line_nom_add.setClearButtonEnabled(False)
            self.line_nom_add.setStyleSheet(styleNormal)
            s2=True
        if(prenom==""):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Prénom\t\n*Champ Obligatoire")
            msgBox.exec()
            self.line_prenom_add.setClearButtonEnabled(True)
            self.line_prenom_add.setStyleSheet(styleErreur)
            s3=False
        else:
            self.line_prenom_add.setClearButtonEnabled(False)
            self.line_prenom_add.setStyleSheet(styleNormal)
            s3=True
        if(age=="" or int(age)<18):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Age\t\n*Champ Obligatoire\n*Age > 18 Ans")
            msgBox.exec()
            self.line_age_add.setClearButtonEnabled(True)
            self.line_age_add.setStyleSheet(styleErreur)
            s4=False
        else:
            self.line_age_add.setClearButtonEnabled(False)
            self.line_age_add.setStyleSheet(styleNormal)
            s4=True
        if(adresse==""):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Adresse\t\n*Champ Obligatoire\n")
            msgBox.exec()
            self.line_adresse_add.setClearButtonEnabled(True)
            self.line_adresse_add.setStyleSheet(styleErreur)
            s5=False
        else:
            self.line_adresse_add.setClearButtonEnabled(False)
            self.line_adresse_add.setStyleSheet(styleNormal)
            s5=True
        if(mail=="" or not re.fullmatch(regexmail, mail)):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Mail\t\n*Champ Obligatoire\n*Sous La forme: User@domaine.com")
            msgBox.exec()
            self.line_mail_add.setClearButtonEnabled(True)
            self.line_mail_add.setStyleSheet(styleErreur)
            s6=False
        else:
            self.line_mail_add.setClearButtonEnabled(False)
            self.line_mail_add.setStyleSheet(styleNormal)
            s6=True
        if(tel=="" or len(tel)<8):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Numéro Tél.\t\n*Champ Obligatoire\n*8 Chiffres")
            msgBox.exec()
            self.line_numero_add.setClearButtonEnabled(True)
            self.line_numero_add.setStyleSheet(styleErreur)
            s7=False
        else:
            self.line_numero_add.setClearButtonEnabled(False)
            self.line_numero_add.setStyleSheet(styleNormal)
            s7=True
        if(s1 and s2 and s3 and s4 and s5 and s6 and s7):
            self.saveClient=False
            self.Data.add_client(cin,nom,prenom,age,adresse,mail,tel)
            msg=MessageBoxValid()
            msg.setText("Information: \n\nMr/Mrs "+str(nom)+" "+str(prenom)+" est Ajoutée Avec Succées")
            msg.exec()
            self.click_btn_add_client()
        
    #function btn_delete_client_confirm
    def click_btn_delete_client_confirm(self):
        styleNormal="border-radius:10px;border: 2px solid rgb(133, 183, 198);background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        styleErreur="border-radius:10px;border: 2px solid red;background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        msgBox=MessageBoxErreur()
        cin=self.line_cin_delete.text()
        if(cin=="" or len(cin)<8 or cin not in self.Data.dictClient.keys()):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ CIN\t\n*Champs Obligatoire\n*8 Chiffres\n*Pas de Client Avec Cette Cin\n")
            msgBox.exec()
            self.line_cin_delete.setClearButtonEnabled(True)
            self.line_cin_delete.setStyleSheet(styleErreur)
        else:
            self.line_cin_delete.setClearButtonEnabled(False)
            self.line_cin_delete.setStyleSheet(styleNormal)
            msg=MessageBoxValid()
            msg.setText("Information: \n\nLe Client "+str((self.Data.dictClient[cin]).nom)+" "+str((self.Data.dictClient[cin]).prenom)+" est Supprrimé avec Succé")
            msg.exec()
            self.Data.delete_client(cin)
            self.click_btn_delete_client()
            self.saveClient=False
    
    #function btn_update_adresse
    def click_btn_update_adresse(self):
        s1,s2,=True,True
        msgBox=MessageBoxErreur()
        styleNormal="border-radius:10px;border: 2px solid rgb(133, 183, 198);background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        styleErreur="border-radius:10px;border: 2px solid red;background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        cin=self.line_cin_adresse.text()
        adresse=self.line_new_adresse.text()
        if(cin=="" or len(cin)<8 or cin not in self.Data.dictClient):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ CIN\t\n*Champs Obligatoire\n*8 Chiffres\n*Pas de Client Avec Cette Cin\n")
            msgBox.exec()
            self.line_cin_adresse.setClearButtonEnabled(True)
            self.line_cin_adresse.setStyleSheet(styleErreur)
            s1=False
        else:
            self.line_cin_adresse.setClearButtonEnabled(False)
            self.line_cin_adresse.setStyleSheet(styleNormal)
            s1=True
        if(adresse==""):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Adresse\t\n*Champ Obligatoire")
            msgBox.exec()
            self.line_new_adresse.setClearButtonEnabled(True)
            self.line_new_adresse.setStyleSheet(styleErreur)
            s2=False
        else:
            self.line_new_adresse.setClearButtonEnabled(False)
            self.line_new_adresse.setStyleSheet(styleNormal)
            s2=True
        if(s1 and s2):
            self.saveClient=False
            msg=MessageBoxValid()
            self.Data.update_client_Adresse(cin,adresse)
            msg.setText("Information: \n\nL'adresse de Mr/Mrs "+str(self.Data.dictClient[cin].nom)+" "+str(self.Data.dictClient[cin].prenom)+" est Modifiée Avec Succé")
            msg.exec()
            self.click_btn_adresse()

    #btn_update_tel
    def click_btn_update_tel(self):
        s1,s2,=True,True
        msgBox=MessageBoxErreur()
        styleNormal="border-radius:10px;border: 2px solid rgb(133, 183, 198);background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        styleErreur="border-radius:10px;border: 2px solid red;background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        cin=self.line_cin_tel.text()
        tel=self.line_new_tel.text()
        if(cin=="" or len(cin)<8 or cin not in self.Data.dictClient):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ CIN\t\n*Champs Obligatoire\n*8 Chiffres\n*Pas de Client Avec Cette Cin\n")
            msgBox.exec()
            self.line_cin_tel.setClearButtonEnabled(True)
            self.line_cin_tel.setStyleSheet(styleErreur)
            s1=False
        else:
            self.line_cin_tel.setClearButtonEnabled(False)
            self.line_cin_tel.setStyleSheet(styleNormal)
            s1=True
        if(tel=="" or len(tel)<8):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Numéro Tél.\t\n*Champ Obligatoire\n*8 chiffres")
            msgBox.exec()
            self.line_new_tel.setClearButtonEnabled(True)
            self.line_new_tel.setStyleSheet(styleErreur)
            s2=False
        else:
            self.line_new_tel.setClearButtonEnabled(False)
            self.line_new_tel.setStyleSheet(styleNormal)
            s2=True
        if(s1 and s2):
            self.saveClient=False
            msg=MessageBoxValid()
            self.Data.update_client_tel(cin,tel)
            msg.setText("Information: \n\nLe Numéro de Mr/Mrs "+str(self.Data.dictClient[cin].nom)+" "+str(self.Data.dictClient[cin].prenom)+" est Modifié Avec Succé")
            msg.exec()
            self.click_btn_tel()

    #function btn_update_mail
    def click_btn_update_mail(self):
        s1,s2,=True,True
        regexmail=r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
        msgBox=MessageBoxErreur()
        styleNormal="border-radius:10px;border: 2px solid rgb(133, 183, 198);background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        styleErreur="border-radius:10px;border: 2px solid red;background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        cin=self.line_cin_mail.text()
        mail=self.line_new_mail.text()
        if(cin=="" or len(cin)<8 or cin not in self.Data.dictClient):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ CIN\t\n*Champs Obligatoire\n*8 Chiffres\n*Pas de Client Avec Cette Cin\n")
            msgBox.exec()
            self.line_cin_mail.setClearButtonEnabled(True)
            self.line_cin_mail.setStyleSheet(styleErreur)
            s1=False
        else:
            self.line_cin_mail.setClearButtonEnabled(False)
            self.line_cin_mail.setStyleSheet(styleNormal)
            s1=True
        if(mail=="" or not re.fullmatch(regexmail, mail)):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Mail\t\n*Champ Obligatoire\n*Sous La forme: User@domaine.com")
            msgBox.exec()
            self.line_new_mail.setClearButtonEnabled(True)
            self.line_new_mail.setStyleSheet(styleErreur)
            s2=False
        else:
            self.line_new_mail.setClearButtonEnabled(False)
            self.line_new_mail.setStyleSheet(styleNormal)
            s2=True
        if(s1 and s2):
            self.saveClient=False
            msg=MessageBoxValid()
            self.Data.update_client_mail(cin,mail)
            msg.setText("Information: \n\nL'adresse Mail de Mr/Mrs "+str(self.Data.dictClient[cin].nom)+" "+str(self.Data.dictClient[cin].prenom)+" est Modifiée Avec Succé")
            msg.exec()
            self.click_btn_mail()
    #function btn_search_with_cin
    def click_btn_search_with_cin(self):
        msgBox=MessageBoxErreur()
        styleNormal="border-radius:10px;border: 2px solid rgb(133, 183, 198);background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        styleErreur="border-radius:10px;border: 2px solid red;background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        cin=self.line_cin_search.text()
        if(cin=="" or len(cin)<8 or cin not in self.Data.dictClient):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ CIN\t\n*Champs Obligatoire\n*8 Chiffres\n*Pas de Client Avec Cette Cin\n")
            msgBox.exec()
            self.line_cin_search.setClearButtonEnabled(True)
            self.line_cin_search.setStyleSheet(styleErreur)
            self.click_btn_search_client()
            self.line_cin_search.setText(cin)
        else:
            self.line_cin_search.setClearButtonEnabled(False)
            self.line_cin_search.setStyleSheet(styleNormal)
            self.Data.search_client_cin(cin,self.line_nom_search,self.line_prenom_search,self.line_age_search,self.line_adresse_search,self.line_Mail_search,self.line_tel_search)



    #function btn_add_client 
    def click_btn_add_client(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_adresse.setVisible(False)
        self.btn_tel.setVisible(False)
        self.btn_mail.setVisible(False)
        self.btn_add_client.setStyleSheet(styleClicked)
        self.btn_delete_client.setStyleSheet(stylenormal)
        self.btn_update_client.setStyleSheet(stylenormal)
        self.btn_adresse.setStyleSheet(stylenormal)
        self.btn_tel.setStyleSheet(stylenormal)
        self.btn_mail.setStyleSheet(stylenormal)
        self.btn_affiche_client.setStyleSheet(stylenormal)
        self.btn_search_client.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(13)
        self.line_Cin_add.setText("")
        self.line_nom_add.setText("")
        self.line_prenom_add.setText("")
        self.line_age_add.setText("")
        self.line_adresse_add.setText("")
        self.line_mail_add.setText("")
        self.line_numero_add.setText("")
    #function btn_delete_client
    def click_btn_delete_client(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_adresse.setVisible(False)
        self.btn_tel.setVisible(False)
        self.btn_mail.setVisible(False)
        self.btn_add_client.setStyleSheet(stylenormal)
        self.btn_delete_client.setStyleSheet(styleClicked)
        self.btn_update_client.setStyleSheet(stylenormal)
        self.btn_adresse.setStyleSheet(stylenormal)
        self.btn_tel.setStyleSheet(stylenormal)
        self.btn_mail.setStyleSheet(stylenormal)
        self.btn_affiche_client.setStyleSheet(stylenormal)
        self.btn_search_client.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(14)
        self.line_cin_delete.setText("")

    #function btn_update_client
    def click_btn_update_client(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_adresse.setVisible(True)
        self.btn_tel.setVisible(True)
        self.btn_mail.setVisible(True)
        self.btn_add_client.setStyleSheet(stylenormal)
        self.btn_delete_client.setStyleSheet(stylenormal)
        self.btn_update_client.setStyleSheet(styleClicked)
        self.btn_adresse.setStyleSheet(styleClicked)
        self.btn_tel.setStyleSheet(stylenormal)
        self.btn_mail.setStyleSheet(stylenormal)
        self.btn_affiche_client.setStyleSheet(stylenormal)
        self.btn_search_client.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(15)
        self.line_cin_adresse.setText("")
        self.line_new_adresse.setText("")

    #function btn_adresse
    def click_btn_adresse(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_add_client.setStyleSheet(stylenormal)
        self.btn_delete_client.setStyleSheet(stylenormal)
        self.btn_adresse.setStyleSheet(styleClicked)
        self.btn_tel.setStyleSheet(stylenormal)
        self.btn_mail.setStyleSheet(stylenormal)
        self.btn_affiche_client.setStyleSheet(stylenormal)
        self.btn_search_client.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(15)
        self.line_cin_adresse.setText("")
        self.line_new_adresse.setText("")
    #function btn_tel
    def click_btn_tel(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_add_client.setStyleSheet(stylenormal)
        self.btn_delete_client.setStyleSheet(stylenormal)
        self.btn_adresse.setStyleSheet(stylenormal)
        self.btn_tel.setStyleSheet(styleClicked)
        self.btn_mail.setStyleSheet(stylenormal)
        self.btn_affiche_client.setStyleSheet(stylenormal)
        self.btn_search_client.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(16)
        self.line_cin_tel.setText("")
        self.line_new_tel.setText("")

    #function btn_mail
    def click_btn_mail(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_add_client.setStyleSheet(stylenormal)
        self.btn_delete_client.setStyleSheet(stylenormal)
        self.btn_adresse.setStyleSheet(stylenormal)
        self.btn_tel.setStyleSheet(stylenormal)
        self.btn_mail.setStyleSheet(styleClicked)
        self.btn_affiche_client.setStyleSheet(stylenormal)
        self.btn_search_client.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(17)
        self.line_cin_mail.setText("")
        self.line_new_mail.setText("")

    #function btn_affiche_client
    def click_btn_affiche_client(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_adresse.setVisible(False)
        self.btn_tel.setVisible(False)
        self.btn_mail.setVisible(False)
        self.btn_add_client.setStyleSheet(stylenormal)
        self.btn_delete_client.setStyleSheet(stylenormal)
        self.btn_update_client.setStyleSheet(stylenormal)
        self.btn_adresse.setStyleSheet(stylenormal)
        self.btn_tel.setStyleSheet(stylenormal)
        self.btn_mail.setStyleSheet(stylenormal)
        self.btn_affiche_client.setStyleSheet(styleClicked)
        self.btn_search_client.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(18)
        self.Data.list_client(self.View_Client)

    #function btn_search_client
    def click_btn_search_client(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_adresse.setVisible(False)
        self.btn_tel.setVisible(False)
        self.btn_mail.setVisible(False)
        self.btn_add_client.setStyleSheet(stylenormal)
        self.btn_delete_client.setStyleSheet(stylenormal)
        self.btn_update_client.setStyleSheet(stylenormal)
        self.btn_adresse.setStyleSheet(stylenormal)
        self.btn_tel.setStyleSheet(stylenormal)
        self.btn_mail.setStyleSheet(stylenormal)
        self.btn_affiche_client.setStyleSheet(stylenormal)
        self.btn_search_client.setStyleSheet(styleClicked)
        self.Main_frame.setCurrentIndex(19)
        self.line_cin_search.setText("")
        self.line_nom_search.setText("")
        self.line_prenom_search.setText("")
        self.line_age_search.setText("")
        self.line_adresse_search.setText("")
        self.line_Mail_search.setText("")
        self.line_tel_search.setText("")
    
    ######################################################### Espace Voiture ############################################################

    #function btn_add_voiture
    def click_btn_add_voiture(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_matricule.setVisible(False)
        self.btn_marque.setVisible(False)
        self.btn_age.setVisible(False)
        self.btn_prix.setVisible(False)
        self.btn_couleur.setVisible(False)
        self.btn_matricule_voiture.setVisible(False)
        self.btn_marque_voiture.setVisible(False)
        self.btn_couleur_voiture.setVisible(False)
        self.btn_disponible_voiture.setVisible(False)
        self.btn_louee_voiture.setVisible(False)
        self.btn_louee_2_voiture.setVisible(False)
        self.btn_add_voiture.setStyleSheet(styleClicked)
        self.btn_delete_voiture.setStyleSheet(stylenormal)
        self.btn_matricule.setStyleSheet(stylenormal)
        self.btn_marque.setStyleSheet(stylenormal)
        self.btn_age.setStyleSheet(stylenormal)
        self.btn_update_voiture.setStyleSheet(stylenormal)
        self.btn_prix.setStyleSheet(stylenormal)
        self.btn_couleur.setStyleSheet(stylenormal)
        self.btn_list_car.setStyleSheet(stylenormal)
        self.btn_search_voiture.setStyleSheet(stylenormal)
        self.btn_matricule_voiture.setStyleSheet(stylenormal)
        self.btn_marque_voiture.setStyleSheet(stylenormal)
        self.btn_couleur_voiture.setStyleSheet(stylenormal)
        self.btn_disponible_voiture.setStyleSheet(stylenormal)
        self.btn_louee_voiture.setStyleSheet(stylenormal)
        self.btn_louee_2_voiture.setStyleSheet(stylenormal)
        self.btn_matricule.setVisible(False)
        self.btn_marque.setVisible(False)
        self.btn_age.setVisible(False)
        self.btn_prix.setVisible(False)
        self.btn_couleur.setVisible(False)
        self.btn_matricule_voiture.setVisible(False)
        self.btn_marque_voiture.setVisible(False)
        self.btn_couleur_voiture.setVisible(False)
        self.btn_disponible_voiture.setVisible(False)
        self.btn_louee_voiture.setVisible(False)
        self.btn_louee_2_voiture.setVisible(False)
        self.Main_frame.setCurrentIndex(0)
        self.line_Matricule.setText("")
        self.line_marque.setText("")
        self.line_couleur.setText("")
        self.line_date.setDate(QDate(2000,1,1))
        self.line_prix.setText("")
        #Mazel ne9sa

    #function btn_delete_voiture
    def click_btn_delete_voiture(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_matricule.setVisible(True)
        self.btn_marque.setVisible(True)
        self.btn_age.setVisible(True)
        self.btn_prix.setVisible(False)
        self.btn_couleur.setVisible(False)
        self.btn_matricule_voiture.setVisible(False)
        self.btn_marque_voiture.setVisible(False)
        self.btn_couleur_voiture.setVisible(False)
        self.btn_disponible_voiture.setVisible(False)
        self.btn_louee_voiture.setVisible(False)
        self.btn_louee_2_voiture.setVisible(False)
        self.btn_add_voiture.setStyleSheet(stylenormal)
        self.btn_delete_voiture.setStyleSheet(styleClicked)
        self.btn_matricule.setStyleSheet(styleClicked)
        self.btn_marque.setStyleSheet(stylenormal)
        self.btn_age.setStyleSheet(stylenormal)
        self.btn_update_voiture.setStyleSheet(stylenormal)
        self.btn_prix.setStyleSheet(stylenormal)
        self.btn_couleur.setStyleSheet(stylenormal)
        self.btn_list_car.setStyleSheet(stylenormal)
        self.btn_search_voiture.setStyleSheet(stylenormal)
        self.btn_matricule_voiture.setStyleSheet(stylenormal)
        self.btn_marque_voiture.setStyleSheet(stylenormal)
        self.btn_couleur_voiture.setStyleSheet(stylenormal)
        self.btn_disponible_voiture.setStyleSheet(stylenormal)
        self.btn_louee_voiture.setStyleSheet(stylenormal)
        self.btn_louee_2_voiture.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(1)
        self.line_Matricule_delete.setText("")

    #function btn_matricule
    def click_btn_matricule(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_matricule.setVisible(True)
        self.btn_marque.setVisible(True)
        self.btn_age.setVisible(True)
        self.btn_prix.setVisible(False)
        self.btn_couleur.setVisible(False)
        self.btn_matricule_voiture.setVisible(False)
        self.btn_marque_voiture.setVisible(False)
        self.btn_couleur_voiture.setVisible(False)
        self.btn_disponible_voiture.setVisible(False)
        self.btn_louee_voiture.setVisible(False)
        self.btn_louee_2_voiture.setVisible(False)
        self.btn_add_voiture.setStyleSheet(stylenormal)
        self.btn_delete_voiture.setStyleSheet(styleClicked)
        self.btn_matricule.setStyleSheet(styleClicked)
        self.btn_marque.setStyleSheet(stylenormal)
        self.btn_age.setStyleSheet(stylenormal)
        self.btn_update_voiture.setStyleSheet(stylenormal)
        self.btn_prix.setStyleSheet(stylenormal)
        self.btn_couleur.setStyleSheet(stylenormal)
        self.btn_list_car.setStyleSheet(stylenormal)
        self.btn_search_voiture.setStyleSheet(stylenormal)
        self.btn_matricule_voiture.setStyleSheet(stylenormal)
        self.btn_marque_voiture.setStyleSheet(stylenormal)
        self.btn_couleur_voiture.setStyleSheet(stylenormal)
        self.btn_disponible_voiture.setStyleSheet(stylenormal)
        self.btn_louee_voiture.setStyleSheet(stylenormal)
        self.btn_louee_2_voiture.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(1)
        self.line_Matricule_delete.setText("")

    #function btn_marque
    def click_btn_marque(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_matricule.setVisible(True)
        self.btn_marque.setVisible(True)
        self.btn_age.setVisible(True)
        self.btn_prix.setVisible(False)
        self.btn_couleur.setVisible(False)
        self.btn_matricule_voiture.setVisible(False)
        self.btn_marque_voiture.setVisible(False)
        self.btn_couleur_voiture.setVisible(False)
        self.btn_disponible_voiture.setVisible(False)
        self.btn_louee_voiture.setVisible(False)
        self.btn_louee_2_voiture.setVisible(False)
        self.btn_add_voiture.setStyleSheet(stylenormal)
        self.btn_delete_voiture.setStyleSheet(styleClicked)
        self.btn_matricule.setStyleSheet(stylenormal)
        self.btn_marque.setStyleSheet(styleClicked)
        self.btn_age.setStyleSheet(stylenormal)
        self.btn_update_voiture.setStyleSheet(stylenormal)
        self.btn_prix.setStyleSheet(stylenormal)
        self.btn_couleur.setStyleSheet(stylenormal)
        self.btn_list_car.setStyleSheet(stylenormal)
        self.btn_search_voiture.setStyleSheet(stylenormal)
        self.btn_matricule_voiture.setStyleSheet(stylenormal)
        self.btn_marque_voiture.setStyleSheet(stylenormal)
        self.btn_couleur_voiture.setStyleSheet(stylenormal)
        self.btn_disponible_voiture.setStyleSheet(stylenormal)
        self.btn_louee_voiture.setStyleSheet(stylenormal)
        self.btn_louee_2_voiture.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(2)
        self.comboBox_delete_marque.clear()
        s=set()
        for i in self.Data.dictVoiture.values():
            s.add(i.marque)
        for i in s:
            self.comboBox_delete_marque.addItem(str(i))
        self.comboBox_delete_marque.setCurrentIndex(-1)


    #function btn_age
    def click_btn_age(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_matricule.setVisible(True)
        self.btn_marque.setVisible(True)
        self.btn_age.setVisible(True)
        self.btn_prix.setVisible(False)
        self.btn_couleur.setVisible(False)
        self.btn_matricule_voiture.setVisible(False)
        self.btn_marque_voiture.setVisible(False)
        self.btn_couleur_voiture.setVisible(False)
        self.btn_disponible_voiture.setVisible(False)
        self.btn_louee_voiture.setVisible(False)
        self.btn_louee_2_voiture.setVisible(False)
        self.btn_add_voiture.setStyleSheet(stylenormal)
        self.btn_delete_voiture.setStyleSheet(styleClicked)
        self.btn_matricule.setStyleSheet(stylenormal)
        self.btn_marque.setStyleSheet(stylenormal)
        self.btn_age.setStyleSheet(styleClicked)
        self.btn_update_voiture.setStyleSheet(stylenormal)
        self.btn_prix.setStyleSheet(stylenormal)
        self.btn_couleur.setStyleSheet(stylenormal)
        self.btn_list_car.setStyleSheet(stylenormal)
        self.btn_search_voiture.setStyleSheet(stylenormal)
        self.btn_matricule_voiture.setStyleSheet(stylenormal)
        self.btn_marque_voiture.setStyleSheet(stylenormal)
        self.btn_couleur_voiture.setStyleSheet(stylenormal)
        self.btn_disponible_voiture.setStyleSheet(stylenormal)
        self.btn_louee_voiture.setStyleSheet(stylenormal)
        self.btn_louee_2_voiture.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(3)
        list_car=self.Data.delete_car_age()
        if(len(list_car)==0):
            self.Attention_label_age.setText("Il n'existe pas des voitures dont leurs age est supérieur à 10 ans,\n Tous Vos voitures sont neuves...")
            self.btn_delete_age_confirm.setEnabled(False)
        else:
            self.Attention_label_age.setText("Il Existe "+str(len(list_car))+" voitures dont leurs age est supérieur à 10 ans,\n Vous êtes sure de les Supprimer ?")
            self.btn_delete_age_confirm.setEnabled(True)
    #function btn_delete_age_confirm
    def click_btn_delete_age_confirm(self):
        msg=MessageBoxValid()
        j=int(date.today().day)
        m=int(date.today().month)
        y=int(date.today().year)
        dateNow=QDate(y,m,j)
        listc=self.Data.delete_car_age()
        l2=[]
        for key,value in self.Data.dictLocation.items():
            if (((self.Data.dictVoiture[value.matricule].date).daysTo(dateNow))>3652.5):
                l2.append(key)
        for i in l2:
            del self.Data.dictLocation[i]
            self.saveLocation=False
        for i in listc:
            del self.Data.dictVoiture[i]
            self.saveVoiture=False
        msg.setText("Information: \n\nSuppression avec Succée\n"+str(len(listc))+" Voitures ont été supprimé\n")
        msg.exec()


    #function btn_update_voiture
    def click_btn_update_voiture(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_matricule.setVisible(False)
        self.btn_marque.setVisible(False)
        self.btn_age.setVisible(False)
        self.btn_prix.setVisible(True)
        self.btn_couleur.setVisible(True)
        self.btn_matricule_voiture.setVisible(False)
        self.btn_marque_voiture.setVisible(False)
        self.btn_couleur_voiture.setVisible(False)
        self.btn_disponible_voiture.setVisible(False)
        self.btn_louee_voiture.setVisible(False)
        self.btn_louee_2_voiture.setVisible(False)
        self.btn_add_voiture.setStyleSheet(stylenormal)
        self.btn_delete_voiture.setStyleSheet(stylenormal)
        self.btn_matricule.setStyleSheet(stylenormal)
        self.btn_marque.setStyleSheet(stylenormal)
        self.btn_age.setStyleSheet(stylenormal)
        self.btn_update_voiture.setStyleSheet(styleClicked)
        self.btn_prix.setStyleSheet(styleClicked)
        self.btn_couleur.setStyleSheet(stylenormal)
        self.btn_list_car.setStyleSheet(stylenormal)
        self.btn_search_voiture.setStyleSheet(stylenormal)
        self.btn_matricule_voiture.setStyleSheet(stylenormal)
        self.btn_marque_voiture.setStyleSheet(stylenormal)
        self.btn_couleur_voiture.setStyleSheet(stylenormal)
        self.btn_disponible_voiture.setStyleSheet(stylenormal)
        self.btn_louee_voiture.setStyleSheet(stylenormal)
        self.btn_louee_2_voiture.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(4)
        self.line_Matricule_update_prix.setText("")
        self.line_new_prix.setText("")
    
    #function btn_prix
    def click_btn_prix(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_matricule.setVisible(False)
        self.btn_marque.setVisible(False)
        self.btn_age.setVisible(False)
        self.btn_prix.setVisible(True)
        self.btn_couleur.setVisible(True)
        self.btn_matricule_voiture.setVisible(False)
        self.btn_marque_voiture.setVisible(False)
        self.btn_couleur_voiture.setVisible(False)
        self.btn_disponible_voiture.setVisible(False)
        self.btn_louee_voiture.setVisible(False)
        self.btn_louee_2_voiture.setVisible(False)
        self.btn_add_voiture.setStyleSheet(stylenormal)
        self.btn_delete_voiture.setStyleSheet(stylenormal)
        self.btn_matricule.setStyleSheet(stylenormal)
        self.btn_marque.setStyleSheet(stylenormal)
        self.btn_age.setStyleSheet(stylenormal)
        self.btn_update_voiture.setStyleSheet(styleClicked)
        self.btn_prix.setStyleSheet(styleClicked)
        self.btn_couleur.setStyleSheet(stylenormal)
        self.btn_list_car.setStyleSheet(stylenormal)
        self.btn_search_voiture.setStyleSheet(stylenormal)
        self.btn_matricule_voiture.setStyleSheet(stylenormal)
        self.btn_marque_voiture.setStyleSheet(stylenormal)
        self.btn_couleur_voiture.setStyleSheet(stylenormal)
        self.btn_disponible_voiture.setStyleSheet(stylenormal)
        self.btn_louee_voiture.setStyleSheet(stylenormal)
        self.btn_louee_2_voiture.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(4)
        self.line_Matricule_update_prix.setText("")
        self.line_new_prix.setText("")
    
    #function btn_couleur
    def click_btn_couleur(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_matricule.setVisible(False)
        self.btn_marque.setVisible(False)
        self.btn_age.setVisible(False)
        self.btn_prix.setVisible(True)
        self.btn_couleur.setVisible(True)
        self.btn_matricule_voiture.setVisible(False)
        self.btn_marque_voiture.setVisible(False)
        self.btn_couleur_voiture.setVisible(False)
        self.btn_disponible_voiture.setVisible(False)
        self.btn_louee_voiture.setVisible(False)
        self.btn_louee_2_voiture.setVisible(False)
        self.btn_add_voiture.setStyleSheet(stylenormal)
        self.btn_delete_voiture.setStyleSheet(stylenormal)
        self.btn_matricule.setStyleSheet(stylenormal)
        self.btn_marque.setStyleSheet(stylenormal)
        self.btn_age.setStyleSheet(stylenormal)
        self.btn_update_voiture.setStyleSheet(styleClicked)
        self.btn_prix.setStyleSheet(stylenormal)
        self.btn_couleur.setStyleSheet(styleClicked)
        self.btn_list_car.setStyleSheet(stylenormal)
        self.btn_search_voiture.setStyleSheet(stylenormal)
        self.btn_matricule_voiture.setStyleSheet(stylenormal)
        self.btn_marque_voiture.setStyleSheet(stylenormal)
        self.btn_couleur_voiture.setStyleSheet(stylenormal)
        self.btn_disponible_voiture.setStyleSheet(stylenormal)
        self.btn_louee_voiture.setStyleSheet(stylenormal)
        self.btn_louee_2_voiture.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(5)
        self.line_Matricule_update_color.setText("")
        self.line_new_color.setText("")

    #function btn_list_car
    def click_btn_list_car(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_matricule.setVisible(False)
        self.btn_marque.setVisible(False)
        self.btn_age.setVisible(False)
        self.btn_prix.setVisible(False)
        self.btn_couleur.setVisible(False)
        self.btn_matricule_voiture.setVisible(False)
        self.btn_marque_voiture.setVisible(False)
        self.btn_couleur_voiture.setVisible(False)
        self.btn_disponible_voiture.setVisible(False)
        self.btn_louee_voiture.setVisible(False)
        self.btn_louee_2_voiture.setVisible(False)
        self.btn_add_voiture.setStyleSheet(stylenormal)
        self.btn_delete_voiture.setStyleSheet(stylenormal)
        self.btn_matricule.setStyleSheet(stylenormal)
        self.btn_marque.setStyleSheet(stylenormal)
        self.btn_age.setStyleSheet(stylenormal)
        self.btn_update_voiture.setStyleSheet(stylenormal)
        self.btn_prix.setStyleSheet(stylenormal)
        self.btn_couleur.setStyleSheet(stylenormal)
        self.btn_list_car.setStyleSheet(styleClicked)
        self.btn_search_voiture.setStyleSheet(stylenormal)
        self.btn_matricule_voiture.setStyleSheet(stylenormal)
        self.btn_marque_voiture.setStyleSheet(stylenormal)
        self.btn_couleur_voiture.setStyleSheet(stylenormal)
        self.btn_disponible_voiture.setStyleSheet(stylenormal)
        self.btn_louee_voiture.setStyleSheet(stylenormal)
        self.btn_louee_2_voiture.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(6)
        self.Data.list_car(self.View_car)
    
    #btn_update_price_confirm
    def click_btn_update_price_confirm(self):
        msgBox=MessageBoxErreur()
        regexmat=r'\b[0-9][0-9][0-9]TN[0-9][0-9][0-9][0-9]\b'
        styleNormal="border-radius:10px;border: 2px solid rgb(133, 183, 198);background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        styleErreur="border-radius:10px;border: 2px solid red;background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        matricule=self.line_Matricule_update_prix.text().upper()
        prix=self.line_new_prix.text()
        regexmat=r'\b[0-9][0-9][0-9]TN[0-9][0-9][0-9][0-9]\b'
        if(not re.fullmatch(regexmat,matricule) or matricule=="" or matricule not in self.Data.dictVoiture.keys()):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Matricule\t\n*SOUS LA FORME: 000 TN 0000\n *Pas De Voiture Avec Cette Matricule\n")
            msgBox.exec()
            self.line_Matricule_update_prix.setClearButtonEnabled(True)
            self.line_Matricule_update_prix.setStyleSheet(styleErreur)
            s1=False
        else:
            self.line_Matricule_update_prix.setClearButtonEnabled(False)
            self.line_Matricule_update_prix.setStyleSheet(styleNormal)
            s1=True
        if(prix=="" or not prix.isdigit()):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Prix\t\n*Champ Obligatoire\n*Doit être en chiffre")
            msgBox.exec()
            self.line_new_prix.setClearButtonEnabled(True)
            self.line_new_prix.setStyleSheet(styleErreur)
            s2=False
        else:
            self.line_new_prix.setClearButtonEnabled(False)
            self.line_new_prix.setStyleSheet(styleNormal)
            s2=True
        if(s1 and s2):
            self.Data.update_car_price(matricule,prix)
            msg=MessageBoxValid()
            self.click_btn_prix()
            self.saveVoiture=False
            msg.setText("Information: \n\n Le prix de la Voiture "+str(matricule)+" est Modifié Avec Succées")
            msg.exec()

    #function btn_update_color_confirm
    def click_btn_update_color_confirm(self):
        msgBox=MessageBoxErreur()
        regexmat=r'\b[0-9][0-9][0-9]TN[0-9][0-9][0-9][0-9]\b'
        styleNormal="border-radius:10px;border: 2px solid rgb(133, 183, 198);background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        styleErreur="border-radius:10px;border: 2px solid red;background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        matricule=self.line_Matricule_update_color.text().upper()
        color=self.line_new_color.text()
        regexmat=r'\b[0-9][0-9][0-9]TN[0-9][0-9][0-9][0-9]\b'
        if(not re.fullmatch(regexmat,matricule) or matricule=="" or matricule not in self.Data.dictVoiture.keys()):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Matricule\t\n*SOUS LA FORME: 000 TN 0000\n *Pas De Voiture Avec Cette Matricule\n")
            msgBox.exec()
            self.line_Matricule_update_color.setClearButtonEnabled(True)
            self.line_Matricule_update_color.setStyleSheet(styleErreur)
            s1=False
        else:
            self.line_Matricule_update_color.setClearButtonEnabled(False)
            self.line_Matricule_update_color.setStyleSheet(styleNormal)
            s1=True
        if(color==""):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Prix\t\n*Champ Obligatoire\n")
            msgBox.exec()
            self.line_new_color.setClearButtonEnabled(True)
            self.line_new_color.setStyleSheet(styleErreur)
            s2=False
        else:
            self.line_new_color.setClearButtonEnabled(False)
            self.line_new_color.setStyleSheet(styleNormal)
            s2=True
        if(s1 and s2):
            self.Data.update_car_color(matricule,color)
            self.click_btn_couleur()
            msg=MessageBoxValid()
            self.saveVoiture=False
            msg.setText("Information: \n\n La couleur de la Voiture "+str(matricule)+" est Modifié Avec Succées")
            msg.exec()


    #function btn_search_voiture
    def click_btn_search_voiture(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_matricule.setVisible(False)
        self.btn_marque.setVisible(False)
        self.btn_age.setVisible(False)
        self.btn_prix.setVisible(False)
        self.btn_couleur.setVisible(False)
        self.btn_matricule_voiture.setVisible(True)
        self.btn_marque_voiture.setVisible(True)
        self.btn_couleur_voiture.setVisible(True)
        self.btn_disponible_voiture.setVisible(True)
        self.btn_louee_voiture.setVisible(True)
        self.btn_louee_2_voiture.setVisible(True)
        self.btn_add_voiture.setStyleSheet(stylenormal)
        self.btn_delete_voiture.setStyleSheet(stylenormal)
        self.btn_matricule.setStyleSheet(stylenormal)
        self.btn_marque.setStyleSheet(stylenormal)
        self.btn_age.setStyleSheet(stylenormal)
        self.btn_update_voiture.setStyleSheet(stylenormal)
        self.btn_prix.setStyleSheet(stylenormal)
        self.btn_couleur.setStyleSheet(stylenormal)
        self.btn_list_car.setStyleSheet(stylenormal)
        self.btn_search_voiture.setStyleSheet(styleClicked)
        self.btn_matricule_voiture.setStyleSheet(styleClicked)
        self.btn_marque_voiture.setStyleSheet(stylenormal)
        self.btn_couleur_voiture.setStyleSheet(stylenormal)
        self.btn_disponible_voiture.setStyleSheet(stylenormal)
        self.btn_louee_voiture.setStyleSheet(stylenormal)
        self.btn_louee_2_voiture.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(7)
        self.line_Matricule_search.setText("")
        self.line_marque_search.setText("")
        self.line_date_search.setDate(QDate(2000,1,1))
        self.line_prix_search.setText("")
        self.line_etat_search.setText("")

    #function btn_matricule_voiture
    def click_btn_matricule_voiture(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_matricule.setVisible(False)
        self.btn_marque.setVisible(False)
        self.btn_age.setVisible(False)
        self.btn_prix.setVisible(False)
        self.btn_couleur.setVisible(False)
        self.btn_matricule_voiture.setVisible(True)
        self.btn_marque_voiture.setVisible(True)
        self.btn_couleur_voiture.setVisible(True)
        self.btn_disponible_voiture.setVisible(True)
        self.btn_louee_voiture.setVisible(True)
        self.btn_louee_2_voiture.setVisible(True)
        self.btn_add_voiture.setStyleSheet(stylenormal)
        self.btn_delete_voiture.setStyleSheet(stylenormal)
        self.btn_matricule.setStyleSheet(stylenormal)
        self.btn_marque.setStyleSheet(stylenormal)
        self.btn_age.setStyleSheet(stylenormal)
        self.btn_update_voiture.setStyleSheet(stylenormal)
        self.btn_prix.setStyleSheet(stylenormal)
        self.btn_couleur.setStyleSheet(stylenormal)
        self.btn_list_car.setStyleSheet(stylenormal)
        self.btn_search_voiture.setStyleSheet(styleClicked)
        self.btn_matricule_voiture.setStyleSheet(styleClicked)
        self.btn_marque_voiture.setStyleSheet(stylenormal)
        self.btn_couleur_voiture.setStyleSheet(stylenormal)
        self.btn_disponible_voiture.setStyleSheet(stylenormal)
        self.btn_louee_voiture.setStyleSheet(stylenormal)
        self.btn_louee_2_voiture.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(7)
        self.line_Matricule_search.setText("")
        self.line_marque_search.setText("")
        self.line_date_search.setDate(QDate(2000,1,1))
        self.line_prix_search.setText("")
        self.line_etat_search.setText("")
    
    #function btn_marque_voiture
    def click_btn_marque_voiture(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_matricule.setVisible(False)
        self.btn_marque.setVisible(False)
        self.btn_age.setVisible(False)
        self.btn_prix.setVisible(False)
        self.btn_couleur.setVisible(False)
        self.btn_matricule_voiture.setVisible(True)
        self.btn_marque_voiture.setVisible(True)
        self.btn_couleur_voiture.setVisible(True)
        self.btn_disponible_voiture.setVisible(True)
        self.btn_louee_voiture.setVisible(True)
        self.btn_louee_2_voiture.setVisible(True)
        self.btn_add_voiture.setStyleSheet(stylenormal)
        self.btn_delete_voiture.setStyleSheet(stylenormal)
        self.btn_matricule.setStyleSheet(stylenormal)
        self.btn_marque.setStyleSheet(stylenormal)
        self.btn_age.setStyleSheet(stylenormal)
        self.btn_update_voiture.setStyleSheet(stylenormal)
        self.btn_prix.setStyleSheet(stylenormal)
        self.btn_couleur.setStyleSheet(stylenormal)
        self.btn_list_car.setStyleSheet(stylenormal)
        self.btn_search_voiture.setStyleSheet(styleClicked)
        self.btn_matricule_voiture.setStyleSheet(stylenormal)
        self.btn_marque_voiture.setStyleSheet(styleClicked)
        self.btn_couleur_voiture.setStyleSheet(stylenormal)
        self.btn_disponible_voiture.setStyleSheet(stylenormal)
        self.btn_louee_voiture.setStyleSheet(stylenormal)
        self.btn_louee_2_voiture.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(8)
        self.comboBox_search_marque.clear()
        self.View_car_marque.setRowCount(0)
        s=set()
        for i in self.Data.dictVoiture.values():
            s.add(i.marque)
        for i in s:
            self.comboBox_search_marque.addItem(str(i))
        self.comboBox_search_marque.setCurrentIndex(-1)

    #function btn_couleur_voiture
    def click_btn_couleur_voiture(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_matricule.setVisible(False)
        self.btn_marque.setVisible(False)
        self.btn_age.setVisible(False)
        self.btn_prix.setVisible(False)
        self.btn_couleur.setVisible(False)
        self.btn_matricule_voiture.setVisible(True)
        self.btn_marque_voiture.setVisible(True)
        self.btn_couleur_voiture.setVisible(True)
        self.btn_disponible_voiture.setVisible(True)
        self.btn_louee_voiture.setVisible(True)
        self.btn_louee_2_voiture.setVisible(True)
        self.btn_add_voiture.setStyleSheet(stylenormal)
        self.btn_delete_voiture.setStyleSheet(stylenormal)
        self.btn_matricule.setStyleSheet(stylenormal)
        self.btn_marque.setStyleSheet(stylenormal)
        self.btn_age.setStyleSheet(stylenormal)
        self.btn_update_voiture.setStyleSheet(stylenormal)
        self.btn_prix.setStyleSheet(stylenormal)
        self.btn_couleur.setStyleSheet(stylenormal)
        self.btn_list_car.setStyleSheet(stylenormal)
        self.btn_search_voiture.setStyleSheet(styleClicked)
        self.btn_matricule_voiture.setStyleSheet(stylenormal)
        self.btn_marque_voiture.setStyleSheet(stylenormal)
        self.btn_couleur_voiture.setStyleSheet(styleClicked)
        self.btn_disponible_voiture.setStyleSheet(stylenormal)
        self.btn_louee_voiture.setStyleSheet(stylenormal)
        self.btn_louee_2_voiture.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(9)
        self.comboBox_search_color.clear()
        self.View_car_color.setRowCount(0)
        s=set()
        for i in self.Data.dictVoiture.values():
            s.add(i.couleur)
        for i in s:
            self.comboBox_search_color.addItem(str(i))
        self.comboBox_search_color.setCurrentIndex(-1)

    #function btn_disponible_voiture
    def click_btn_disponible_voiture(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_matricule.setVisible(False)
        self.btn_marque.setVisible(False)
        self.btn_age.setVisible(False)
        self.btn_prix.setVisible(False)
        self.btn_couleur.setVisible(False)
        self.btn_matricule_voiture.setVisible(True)
        self.btn_marque_voiture.setVisible(True)
        self.btn_couleur_voiture.setVisible(True)
        self.btn_disponible_voiture.setVisible(True)
        self.btn_louee_voiture.setVisible(True)
        self.btn_louee_2_voiture.setVisible(True)
        self.btn_add_voiture.setStyleSheet(stylenormal)
        self.btn_delete_voiture.setStyleSheet(stylenormal)
        self.btn_matricule.setStyleSheet(stylenormal)
        self.btn_marque.setStyleSheet(stylenormal)
        self.btn_age.setStyleSheet(stylenormal)
        self.btn_update_voiture.setStyleSheet(stylenormal)
        self.btn_prix.setStyleSheet(stylenormal)
        self.btn_couleur.setStyleSheet(stylenormal)
        self.btn_list_car.setStyleSheet(stylenormal)
        self.btn_search_voiture.setStyleSheet(styleClicked)
        self.btn_matricule_voiture.setStyleSheet(stylenormal)
        self.btn_marque_voiture.setStyleSheet(stylenormal)
        self.btn_couleur_voiture.setStyleSheet(stylenormal)
        self.btn_disponible_voiture.setStyleSheet(styleClicked)
        self.btn_louee_voiture.setStyleSheet(stylenormal)
        self.btn_louee_2_voiture.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(10)
        self.Data.search_car_dispo(self.View_car_dispo)

    #function btn_louee_voiture
    def click_btn_louee_voiture(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_matricule.setVisible(False)
        self.btn_marque.setVisible(False)
        self.btn_age.setVisible(False)
        self.btn_prix.setVisible(False)
        self.btn_couleur.setVisible(False)
        self.btn_matricule_voiture.setVisible(True)
        self.btn_marque_voiture.setVisible(True)
        self.btn_couleur_voiture.setVisible(True)
        self.btn_disponible_voiture.setVisible(True)
        self.btn_louee_voiture.setVisible(True)
        self.btn_louee_2_voiture.setVisible(True)
        self.btn_add_voiture.setStyleSheet(stylenormal)
        self.btn_delete_voiture.setStyleSheet(stylenormal)
        self.btn_matricule.setStyleSheet(stylenormal)
        self.btn_marque.setStyleSheet(stylenormal)
        self.btn_age.setStyleSheet(stylenormal)
        self.btn_update_voiture.setStyleSheet(stylenormal)
        self.btn_prix.setStyleSheet(stylenormal)
        self.btn_couleur.setStyleSheet(stylenormal)
        self.btn_list_car.setStyleSheet(stylenormal)
        self.btn_search_voiture.setStyleSheet(styleClicked)
        self.btn_matricule_voiture.setStyleSheet(stylenormal)
        self.btn_marque_voiture.setStyleSheet(stylenormal)
        self.btn_couleur_voiture.setStyleSheet(stylenormal)
        self.btn_disponible_voiture.setStyleSheet(stylenormal)
        self.btn_louee_voiture.setStyleSheet(styleClicked)
        self.btn_louee_2_voiture.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(11)
        self.Data.search_car_louee(self.View_car_louee)
    
    #function btn_louee_2_voiture
    def click_btn_louee_2_voiture(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_matricule.setVisible(False)
        self.btn_marque.setVisible(False)
        self.btn_age.setVisible(False)
        self.btn_prix.setVisible(False)
        self.btn_couleur.setVisible(False)
        self.btn_matricule_voiture.setVisible(True)
        self.btn_marque_voiture.setVisible(True)
        self.btn_couleur_voiture.setVisible(True)
        self.btn_disponible_voiture.setVisible(True)
        self.btn_louee_voiture.setVisible(True)
        self.btn_louee_2_voiture.setVisible(True)
        self.btn_add_voiture.setStyleSheet(stylenormal)
        self.btn_delete_voiture.setStyleSheet(stylenormal)
        self.btn_matricule.setStyleSheet(stylenormal)
        self.btn_marque.setStyleSheet(stylenormal)
        self.btn_age.setStyleSheet(stylenormal)
        self.btn_update_voiture.setStyleSheet(stylenormal)
        self.btn_prix.setStyleSheet(stylenormal)
        self.btn_couleur.setStyleSheet(stylenormal)
        self.btn_list_car.setStyleSheet(stylenormal)
        self.btn_search_voiture.setStyleSheet(styleClicked)
        self.btn_matricule_voiture.setStyleSheet(stylenormal)
        self.btn_marque_voiture.setStyleSheet(stylenormal)
        self.btn_couleur_voiture.setStyleSheet(stylenormal)
        self.btn_disponible_voiture.setStyleSheet(stylenormal)
        self.btn_louee_voiture.setStyleSheet(stylenormal)
        self.btn_louee_2_voiture.setStyleSheet(styleClicked)
        self.Main_frame.setCurrentIndex(12)
        self.date_1_search.setDate(QDate(2000,1,1))
        self.date_2_search.setDate(QDate(2000,1,1))
        self.View_car_loue.setRowCount(0)
        
    
    #function btn_add_location
    def click_btn_add_location(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_date.setVisible(False)
        self.btn_duree.setVisible(False)
        self.btn_cin_loc.setVisible(False)
        self.btn_matricule_loc.setVisible(False)
        self.btn_date_loc.setVisible(False)
        self.btn_duree_loc.setVisible(False)
        self.btn_2_date.setVisible(False)
        self.btn_add_location.setStyleSheet(styleClicked)
        self.btn_delete_location.setStyleSheet(stylenormal)
        self.btn_update_location.setStyleSheet(stylenormal)
        self.btn_date.setStyleSheet(stylenormal)
        self.btn_duree.setStyleSheet(stylenormal)
        self.btn_afficher_location.setStyleSheet(stylenormal)
        self.btn_recherche_location.setStyleSheet(stylenormal)
        self.btn_cin_loc.setStyleSheet(stylenormal)
        self.btn_matricule_loc.setStyleSheet(stylenormal)
        self.btn_date_loc.setStyleSheet(stylenormal)
        self.btn_duree_loc.setStyleSheet(stylenormal)
        self.btn_2_date.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(20)
        self.line_num_add_loc.setText("")
        self.line_cin_add_loc.setText("")
        self.dateTime_add_loc.setDate(QDate(2000,1,1))
        self.dateTime_add_loc.setTime(QTime(0,0,0,0))
        self.line_dure_add_loc.setText("")
        num='#'+str(random.randint(0,99999))
        while(True):
            if num in self.Data.dictLocation:
                num='#'+str(random.randint(0,99999))
            else:
                break
        self.line_num_add_loc.setText(str(num))
        self.btn_add_client_location.setVisible(False)


    #function btn_delete_location
    def click_btn_delete_location(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_date.setVisible(False)
        self.btn_duree.setVisible(False)
        self.btn_cin_loc.setVisible(False)
        self.btn_matricule_loc.setVisible(False)
        self.btn_date_loc.setVisible(False)
        self.btn_duree_loc.setVisible(False)
        self.btn_2_date.setVisible(False)
        self.btn_add_location.setStyleSheet(stylenormal)
        self.btn_delete_location.setStyleSheet(styleClicked)
        self.btn_update_location.setStyleSheet(stylenormal)
        self.btn_date.setStyleSheet(stylenormal)
        self.btn_duree.setStyleSheet(stylenormal)
        self.btn_afficher_location.setStyleSheet(stylenormal)
        self.btn_recherche_location.setStyleSheet(stylenormal)
        self.btn_cin_loc.setStyleSheet(stylenormal)
        self.btn_matricule_loc.setStyleSheet(stylenormal)
        self.btn_date_loc.setStyleSheet(stylenormal)
        self.btn_duree_loc.setStyleSheet(stylenormal)
        self.btn_2_date.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(22)
        self.line_numero_del_loc.setText("")
    #function btn_update_location
    def click_btn_update_location(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_date.setVisible(True)
        self.btn_duree.setVisible(True)
        self.btn_cin_loc.setVisible(False)
        self.btn_matricule_loc.setVisible(False)
        self.btn_date_loc.setVisible(False)
        self.btn_duree_loc.setVisible(False)
        self.btn_2_date.setVisible(False)
        self.btn_add_location.setStyleSheet(stylenormal)
        self.btn_delete_location.setStyleSheet(stylenormal)
        self.btn_update_location.setStyleSheet(styleClicked)
        self.btn_date.setStyleSheet(styleClicked)
        self.btn_duree.setStyleSheet(stylenormal)
        self.btn_afficher_location.setStyleSheet(stylenormal)
        self.btn_recherche_location.setStyleSheet(stylenormal)
        self.btn_cin_loc.setStyleSheet(stylenormal)
        self.btn_matricule_loc.setStyleSheet(stylenormal)
        self.btn_date_loc.setStyleSheet(stylenormal)
        self.btn_duree_loc.setStyleSheet(stylenormal)
        self.btn_2_date.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(23)
        self.line_numero_update_loc.setText("")
        self.dateTime_new_date_loc.setDate(QDate(2000,1,1))
        self.dateTime_new_date_loc.setTime(QTime(0,0,0,0))

    #function btn_date
    def click_btn_date(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_cin_loc.setVisible(False)
        self.btn_matricule_loc.setVisible(False)
        self.btn_date_loc.setVisible(False)
        self.btn_duree_loc.setVisible(False)
        self.btn_2_date.setVisible(False)
        self.btn_add_location.setStyleSheet(stylenormal)
        self.btn_delete_location.setStyleSheet(stylenormal)
        self.btn_update_location.setStyleSheet(styleClicked)
        self.btn_date.setStyleSheet(styleClicked)
        self.btn_duree.setStyleSheet(stylenormal)
        self.btn_afficher_location.setStyleSheet(stylenormal)
        self.btn_recherche_location.setStyleSheet(stylenormal)
        self.btn_cin_loc.setStyleSheet(stylenormal)
        self.btn_matricule_loc.setStyleSheet(stylenormal)
        self.btn_date_loc.setStyleSheet(stylenormal)
        self.btn_duree_loc.setStyleSheet(stylenormal)
        self.btn_2_date.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(23)
        self.line_numero_update_loc.setText("")
        self.dateTime_new_date_loc.setDate(QDate(2000,1,1))
        self.dateTime_new_date_loc.setTime(QTime(0,0,0,0))
    
    #function btn_duree
    def click_btn_duree(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_cin_loc.setVisible(False)
        self.btn_matricule_loc.setVisible(False)
        self.btn_date_loc.setVisible(False)
        self.btn_duree_loc.setVisible(False)
        self.btn_2_date.setVisible(False)
        self.btn_add_location.setStyleSheet(stylenormal)
        self.btn_delete_location.setStyleSheet(stylenormal)
        self.btn_update_location.setStyleSheet(styleClicked)
        self.btn_date.setStyleSheet(stylenormal)
        self.btn_duree.setStyleSheet(styleClicked)
        self.btn_afficher_location.setStyleSheet(stylenormal)
        self.btn_recherche_location.setStyleSheet(stylenormal)
        self.btn_cin_loc.setStyleSheet(stylenormal)
        self.btn_matricule_loc.setStyleSheet(stylenormal)
        self.btn_date_loc.setStyleSheet(stylenormal)
        self.btn_duree_loc.setStyleSheet(stylenormal)
        self.btn_2_date.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(24)
        self.line_numero_loc_date.setText("")
        self.line_new_duree_loc.setText("")

    #function btn_afficher_location
    def click_btn_afficher_location(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_date.setVisible(False)
        self.btn_duree.setVisible(False)
        self.btn_cin_loc.setVisible(False)
        self.btn_matricule_loc.setVisible(False)
        self.btn_date_loc.setVisible(False)
        self.btn_duree_loc.setVisible(False)
        self.btn_2_date.setVisible(False)
        self.btn_add_location.setStyleSheet(stylenormal)
        self.btn_delete_location.setStyleSheet(stylenormal)
        self.btn_update_location.setStyleSheet(stylenormal)
        self.btn_date.setStyleSheet(stylenormal)
        self.btn_duree.setStyleSheet(stylenormal)
        self.btn_afficher_location.setStyleSheet(styleClicked)
        self.btn_recherche_location.setStyleSheet(stylenormal)
        self.btn_cin_loc.setStyleSheet(stylenormal)
        self.btn_matricule_loc.setStyleSheet(stylenormal)
        self.btn_date_loc.setStyleSheet(stylenormal)
        self.btn_duree_loc.setStyleSheet(stylenormal)
        self.btn_2_date.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(25)
        self.Data.list_loc(self.View_location)

    #function btn_recherche_location
    def click_btn_recherche_location(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_date.setVisible(False)
        self.btn_duree.setVisible(False)
        self.btn_cin_loc.setVisible(True)
        self.btn_matricule_loc.setVisible(True)
        self.btn_date_loc.setVisible(True)
        self.btn_duree_loc.setVisible(True)
        self.btn_2_date.setVisible(True)
        self.btn_add_location.setStyleSheet(stylenormal)
        self.btn_delete_location.setStyleSheet(stylenormal)
        self.btn_update_location.setStyleSheet(stylenormal)
        self.btn_date.setStyleSheet(stylenormal)
        self.btn_duree.setStyleSheet(stylenormal)
        self.btn_afficher_location.setStyleSheet(stylenormal)
        self.btn_recherche_location.setStyleSheet(styleClicked)
        self.btn_cin_loc.setStyleSheet(styleClicked)
        self.btn_matricule_loc.setStyleSheet(stylenormal)
        self.btn_date_loc.setStyleSheet(stylenormal)
        self.btn_duree_loc.setStyleSheet(stylenormal)
        self.btn_2_date.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(26)
        self.line_cin_search_loc.setText("")
        self.View_loc_cin.setRowCount(0)

    #function btn_matricule_loc
    def click_btn_matricule_loc(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_date.setVisible(False)
        self.btn_duree.setVisible(False)
        self.btn_cin_loc.setVisible(True)
        self.btn_matricule_loc.setVisible(True)
        self.btn_date_loc.setVisible(True)
        self.btn_duree_loc.setVisible(True)
        self.btn_2_date.setVisible(True)
        self.btn_add_location.setStyleSheet(stylenormal)
        self.btn_delete_location.setStyleSheet(stylenormal)
        self.btn_update_location.setStyleSheet(stylenormal)
        self.btn_date.setStyleSheet(stylenormal)
        self.btn_duree.setStyleSheet(stylenormal)
        self.btn_afficher_location.setStyleSheet(stylenormal)
        self.btn_recherche_location.setStyleSheet(styleClicked)
        self.btn_cin_loc.setStyleSheet(stylenormal)
        self.btn_matricule_loc.setStyleSheet(styleClicked)
        self.btn_date_loc.setStyleSheet(stylenormal)
        self.btn_duree_loc.setStyleSheet(stylenormal)
        self.btn_2_date.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(27)
        self.line_Matricule_search_loc.setText("")
        self.View_loc_matricule.setRowCount(0)
        

    #function btn_date_loc
    def click_btn_date_loc(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_date.setVisible(False)
        self.btn_duree.setVisible(False)
        self.btn_add_location.setStyleSheet(stylenormal)
        self.btn_delete_location.setStyleSheet(stylenormal)
        self.btn_update_location.setStyleSheet(stylenormal)
        self.btn_date.setStyleSheet(stylenormal)
        self.btn_duree.setStyleSheet(stylenormal)
        self.btn_afficher_location.setStyleSheet(stylenormal)
        self.btn_recherche_location.setStyleSheet(styleClicked)
        self.btn_cin_loc.setStyleSheet(stylenormal)
        self.btn_matricule_loc.setStyleSheet(stylenormal)
        self.btn_date_loc.setStyleSheet(styleClicked)
        self.btn_duree_loc.setStyleSheet(stylenormal)
        self.btn_2_date.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(28)
        self.dateTime_search_loc.setDate(QDate(2000,1,1))
        self.View_loc_date.setRowCount(0)
    
    #function btn_duree_loc
    def click_btn_duree_loc(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_date.setVisible(False)
        self.btn_duree.setVisible(False)
        self.btn_add_location.setStyleSheet(stylenormal)
        self.btn_delete_location.setStyleSheet(stylenormal)
        self.btn_update_location.setStyleSheet(stylenormal)
        self.btn_date.setStyleSheet(stylenormal)
        self.btn_duree.setStyleSheet(stylenormal)
        self.btn_afficher_location.setStyleSheet(stylenormal)
        self.btn_recherche_location.setStyleSheet(styleClicked)
        self.btn_cin_loc.setStyleSheet(stylenormal)
        self.btn_matricule_loc.setStyleSheet(stylenormal)
        self.btn_date_loc.setStyleSheet(stylenormal)
        self.btn_duree_loc.setStyleSheet(styleClicked)
        self.btn_2_date.setStyleSheet(stylenormal)
        self.Main_frame.setCurrentIndex(29)
        self.line_duree_search_loc.setText("")
        self.View_car_dure.setRowCount(0)
    
    #function btn_2_date
    def click_btn_2_date(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_date.setVisible(False)
        self.btn_duree.setVisible(False)
        self.btn_add_location.setStyleSheet(stylenormal)
        self.btn_delete_location.setStyleSheet(stylenormal)
        self.btn_update_location.setStyleSheet(stylenormal)
        self.btn_date.setStyleSheet(stylenormal)
        self.btn_duree.setStyleSheet(stylenormal)
        self.btn_afficher_location.setStyleSheet(stylenormal)
        self.btn_recherche_location.setStyleSheet(styleClicked)
        self.btn_cin_loc.setStyleSheet(stylenormal)
        self.btn_matricule_loc.setStyleSheet(stylenormal)
        self.btn_date_loc.setStyleSheet(stylenormal)
        self.btn_duree_loc.setStyleSheet(stylenormal)
        self.btn_2_date.setStyleSheet(styleClicked)
        self.Main_frame.setCurrentIndex(30)
        self.dateEdit_1_search_loc.setDate(QDate(2000,1,1))
        self.dateEdit_2_search_loc.setDate(QDate(2000,1,1))
        self.View_car_dure.setRowCount(0)
    

    #function to all back button 
    def click_back(self):
        styleClicked="color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";"
        stylenormal="QPushButton{color: rgb(209, 208, 222);background-color: rgb(46, 47, 65);border:0px solid;border-radius:20px;margin:10px;font: 87 12pt ""Segoe UI"";}QPushButton:hover{color: rgb(46, 47, 65);background-color:  rgb(209, 208, 222);}"
        self.btn_date.setVisible(False)
        self.btn_duree.setVisible(False)
        self.btn_adresse.setVisible(False)
        self.btn_tel.setVisible(False)
        self.btn_mail.setVisible(False)
        self.btn_matricule.setVisible(False)
        self.btn_marque.setVisible(False)
        self.btn_age.setVisible(False)
        self.btn_prix.setVisible(False)
        self.btn_couleur.setVisible(False)
        self.btn_matricule_voiture.setVisible(False)
        self.btn_marque_voiture.setVisible(False)
        self.btn_couleur_voiture.setVisible(False)
        self.btn_disponible_voiture.setVisible(False)
        self.btn_louee_voiture.setVisible(False)
        self.btn_louee_2_voiture.setVisible(False)
        self.btn_add_voiture.setStyleSheet(stylenormal)
        self.btn_delete_voiture.setStyleSheet(stylenormal)
        self.btn_matricule.setStyleSheet(stylenormal)
        self.btn_marque.setStyleSheet(stylenormal)
        self.btn_age.setStyleSheet(stylenormal)
        self.btn_update_voiture.setStyleSheet(stylenormal)
        self.btn_prix.setStyleSheet(stylenormal)
        self.btn_couleur.setStyleSheet(stylenormal)
        self.btn_list_car.setStyleSheet(stylenormal)
        self.btn_search_voiture.setStyleSheet(stylenormal)
        self.btn_matricule_voiture.setStyleSheet(stylenormal)
        self.btn_marque_voiture.setStyleSheet(stylenormal)
        self.btn_couleur_voiture.setStyleSheet(stylenormal)
        self.btn_disponible_voiture.setStyleSheet(stylenormal)
        self.btn_louee_voiture.setStyleSheet(stylenormal)
        self.btn_louee_2_voiture.setStyleSheet(stylenormal)
        self.btn_add_client.setStyleSheet(stylenormal)
        self.btn_delete_client.setStyleSheet(stylenormal)
        self.btn_update_client.setStyleSheet(stylenormal)
        self.btn_adresse.setStyleSheet(stylenormal)
        self.btn_tel.setStyleSheet(stylenormal)
        self.btn_mail.setStyleSheet(stylenormal)
        self.btn_affiche_client.setStyleSheet(stylenormal)
        self.btn_search_client.setStyleSheet(stylenormal)
        self.btn_add_location.setStyleSheet(stylenormal)
        self.btn_delete_location.setStyleSheet(stylenormal)
        self.btn_update_location.setStyleSheet(stylenormal)
        self.btn_date.setStyleSheet(stylenormal)
        self.btn_duree.setStyleSheet(stylenormal)
        self.btn_afficher_location.setStyleSheet(stylenormal)
        self.btn_recherche_location.setStyleSheet(stylenormal)
        self.btn_cin_loc.setStyleSheet(stylenormal)
        self.btn_matricule_loc.setStyleSheet(stylenormal)
        self.btn_date_loc.setStyleSheet(stylenormal)
        self.btn_duree_loc.setStyleSheet(stylenormal)
        self.btn_2_date.setStyleSheet(stylenormal)
        self.btn_accueil.setStyleSheet(styleClicked)
        self.Main_frame.setCurrentIndex(31)
        self.stacked_dash.setCurrentIndex(2)

    #function btn_add_car
    def click_btn_add_car(self):
        s1,s2,s3,s4=True,True,True,True
        msgBox=MessageBoxErreur()
        regexmat=r'\b[0-9][0-9][0-9]TN[0-9][0-9][0-9][0-9]\b'
        styleNormal="border-radius:10px;border: 2px solid rgb(133, 183, 198);background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        styleErreur="border-radius:10px;border: 2px solid red;background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        matricule=self.line_Matricule.text().upper()
        marque=self.line_marque.text()
        couleur=self.line_couleur.text()
        date=self.line_date.date()
        prix=self.line_prix.text()
        regexmat=r'\b[0-9][0-9][0-9]TN[0-9][0-9][0-9][0-9]\b'
        if(not re.fullmatch(regexmat,matricule) or matricule=="" or matricule in self.Data.dictVoiture.keys()):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Matricule\t\n*SOUS LA FORME: 000 TN 0000\n *Matricule Doit Etre Unique\n")
            msgBox.exec()
            self.line_Matricule.setClearButtonEnabled(True)
            self.line_Matricule.setStyleSheet(styleErreur)
            s1=False
        else:
            self.line_Matricule.setClearButtonEnabled(False)
            self.line_Matricule.setStyleSheet(styleNormal)
            s1=True
        if(marque==""):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Marque\t\n*Champ Obligatoire")
            msgBox.exec()
            self.line_marque.setClearButtonEnabled(True)
            self.line_marque.setStyleSheet(styleErreur)
            s2=False
        else:
            self.line_marque.setClearButtonEnabled(False)
            self.line_marque.setStyleSheet(styleNormal)
            s2=True
        if(couleur==""):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Couleur\t\n*Champ Obligatoire")
            msgBox.exec()
            self.line_couleur.setClearButtonEnabled(True)
            self.line_couleur.setStyleSheet(styleErreur)
            s3=False
        else:
            self.line_couleur.setClearButtonEnabled(False)
            self.line_couleur.setStyleSheet(styleNormal)
            s3=True
        if(prix=="" or not prix.isdigit()):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Prix\t\n*Champ Obligatoire\n*Doit être en chiffre")
            msgBox.exec()
            self.line_prix.setClearButtonEnabled(True)
            self.line_prix.setStyleSheet(styleErreur)
            s4=False
        else:
            self.line_prix.setClearButtonEnabled(False)
            self.line_prix.setStyleSheet(styleNormal)
            s4=True
        if(s1 and s2 and s3 and s4):
            etat=""
            if(self.radioButton_dispo.isChecked()):
                etat="Disponible"
            else:
                etat="Louée"
            self.saveVoiture=False
            self.click_btn_add_voiture()
            self.Data.add_voiture(matricule,marque,couleur,etat,date,prix)
            msg=MessageBoxValid()
            msg.setText("Information \n\n La Voiture "+str(marque)+" est Ajoutée Avec Succées")
            msg.exec()

    #function btn_delete_car
    def click_btn_delete_car(self):
        matricule=self.line_Matricule_delete.text().upper()
        regexmat=r'\b[0-9][0-9][0-9]TN[0-9][0-9][0-9][0-9]\b'
        styleNormal="border-radius:10px;border: 2px solid rgb(133, 183, 198);background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        styleErreur="border-radius:10px;border: 2px solid red;background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        if(not re.fullmatch(regexmat,matricule) or matricule=="" or matricule not in self.Data.dictVoiture.keys()):
            msgBox=MessageBoxErreur()
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Matricule\t\n*SOUS LA FORME: 000 TN 0000\n *Pas de Voiture Avec Cette Matricule\n")
            msgBox.exec()
            self.line_Matricule_delete.setClearButtonEnabled(True)
            self.line_Matricule_delete.setStyleSheet(styleErreur)
        else:
            self.line_Matricule_delete.setClearButtonEnabled(False)
            self.line_Matricule_delete.setStyleSheet(styleNormal)
            self.saveVoiture=False
            self.click_btn_matricule()
            self.Data.delete_car_mat(matricule)
            msg=MessageBoxValid()
            msg.setText("Information \n\n La Voiture "+str(matricule)+" est Supprimée Avec Succées")
            msg.exec()

    #function btn_delete_car_confirm
    def click_btn_delete_car_confirm(self):
        if(self.comboBox_delete_marque.currentIndex()==-1):
            msgBox=MessageBoxErreur()
            msgBox.setText("Erreur: \n\nS'il Vous plaît choisir la marque des voitures qui vous allez les Supprimer \n\n")
            msgBox.exec()
        else:
            total=self.Data.delete_car_marque(self.comboBox_delete_marque.currentText())
            self.click_btn_marque()
            msg=MessageBoxValid()
            self.saveVoiture=False
            msg.setText("Information: \n\nSuppression Avec Succées\n"+str(total)+" Voitures ont été Supprimer")
            msg.exec()

    #function btn_search_car_matricule_confirm
    def click_btn_search_car_matricule_confirm(self):
        matricule=self.line_Matricule_search.text().upper()
        regexmat=r'\b[0-9][0-9][0-9]TN[0-9][0-9][0-9][0-9]\b'
        styleNormal="border-radius:10px;border: 2px solid rgb(133, 183, 198);background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        styleErreur="border-radius:10px;border: 2px solid red;background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        if((not re.fullmatch(regexmat,matricule)) or matricule=="" or matricule not in self.Data.dictVoiture.keys()):
            msgBox=MessageBoxErreur()
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Matricule\t\n*SOUS LA FORME: 000 TN 0000\n *Pas de Voiture Avec Cette Matricule\n")
            msgBox.exec()
            self.line_Matricule_search.setClearButtonEnabled(True)
            self.line_Matricule_search.setStyleSheet(styleErreur)
        else:
            self.line_Matricule_search.setClearButtonEnabled(False)
            self.line_Matricule_search.setStyleSheet(styleNormal)
            self.Data.search_car_matricule(matricule,self.line_marque_search,self.line_couleur_mat_search,self.line_etat_search,self.line_date_search,self.line_prix_search)
            
    #function btn_search_marque_confirm
    def click_btn_search_marque_confirm(self):
        if(self.comboBox_search_marque.currentIndex()==-1):
            msgBox=MessageBoxErreur()
            msgBox.setText("Erreur: \n\nS'il Vous plaît choisir la marque des voitures\n\n")
            msgBox.exec()
        else:
            self.Data.search_car_marque(self.comboBox_search_marque.currentText(),self.View_car_marque)
    
    #function btn_search_color_confirm
    def click_btn_search_color_confirm(self):
        if(self.comboBox_search_color.currentIndex()==-1):
            msgBox=MessageBoxErreur()
            msgBox.setText("Erreur: \n\nS'il Vous plaît choisir la Couleur des voitures\n\n")
            msgBox.exec()
        else:
            self.Data.search_car_color(self.comboBox_search_color.currentText(),self.View_car_color)

    #function btn_search_loc_dure_confirm
    def click_btn_search_loc_dure_confirm(self):
        styleNormal="border-radius:10px;border: 2px solid rgb(133, 183, 198);background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        styleErreur="border-radius:10px;border: 2px solid red;background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        duree=self.line_duree_search_loc.text()
        if (duree=="" or int(duree)<1):
            msgBox=MessageBoxErreur()
            msgBox.setText("Erreur de saisie: \n\nVérifier le champ durée\n*Champs Obligatoire\n*Doit être supérieur à 1j")
            msgBox.exec()
            self.line_duree_search_loc.setClearButtonEnabled(True)
            self.line_duree_search_loc.setStyleSheet(styleErreur)
        else:
            self.line_duree_search_loc.setClearButtonEnabled(False)
            self.line_duree_search_loc.setStyleSheet(styleNormal)
            self.Data.search_loc_duree(duree,self.View_car_dure)

    #function btn_search_2_date_confirm
    def click_btn_search_2_date_confirm(self):
        date1=self.dateEdit_1_search_loc.date()
        date2=self.dateEdit_2_search_loc.date()
        if(date1>=date2):
            msgBox=MessageBoxErreur()
            msgBox.setText("Erreur de saisie: \n\nVérifier le champ Date 1 ou Date 2\n*Date 1 doit être inférieur à Date 2")
            msgBox.exec()
        else:
            self.Data.search_loc_2_date(date1,date2,self.View_search_2_date)
            
    #function btn_search_car_louee
    def click_btn_search_car_louee(self):
        date1=self.date_1_search.date()
        date2=self.date_2_search.date()
        if(date1>=date2):
            msgBox=MessageBoxErreur()
            msgBox.setText("Erreur de saisie: \n\nVérifier le champ Date 1 ou Date 2\n*Date 1 doit être inférieur à Date 2")
            msgBox.exec()
        else:
            self.Data.search_loc_2_date_car(date1,date2,self.View_car_loue)

    #function btn_update_date_loc_confirm
    def click_btn_update_date_loc_confirm(self):
        numero='#'+self.line_numero_update_loc.text()
        datec=self.dateTime_new_date_loc.dateTime()
        dateNow=QDateTime.currentDateTime()
        styleNormal="border-radius:10px;border: 2px solid rgb(133, 183, 198);background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        styleErreur="border-radius:10px;border: 2px solid red;background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        msgBox=MessageBoxErreur()
        if(numero=="" or numero not in self.Data.dictLocation.keys()):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Numéro\t\n*Champs Obligatoire\n*5 Chiffres\n*Pas de Location avec ce Numéro")
            msgBox.exec()
            self.line_numero_update_loc.setClearButtonEnabled(True)
            self.line_numero_update_loc.setStyleSheet(styleErreur)
            s1=False
        else:
            self.line_numero_update_loc.setClearButtonEnabled(False)
            self.line_numero_update_loc.setStyleSheet(styleNormal)
            s1=True
        if(datec<dateNow):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Date\t\n*Champ Obligatoire\n*Date Doit être supérieur à Date aujourd'hui "+dateNow.toString("dd/MM/yyyy HH'h'"))
            msgBox.exec()
            self.dateTime_new_date_loc.setStyleSheet(styleErreur)
            s2=False
        else:
            self.dateTime_new_date_loc.setStyleSheet(styleNormal)
            s2=True
        if(s1 and s2):
            liste=[]
            s3=True
            car=self.Data.dictLocation[numero].matricule
            for key,item in self.Data.dictLocation.items():
                if(key!=numero and car==item.matricule):
                    liste.append(item)
            for item in liste:
                if(item.date<=datec<=(item.date).addDays(int(item.duree)) or datec<=item.date<=datec.addDays(int(self.Data.dictLocation[numero].duree))):
                    s3=False
                    prop=item
            if(s3==False):
                msgBox.setText("Erreur: \n\nLa voiture "+str(car)+" est déja louée pour cette date à Mr/Mrs "+str(self.Data.dictClient[prop.cin].nom)+" "+str(self.Data.dictClient[prop.cin].prenom))
                msgBox.exec()
                self.dateTime_new_date_loc.setStyleSheet(styleErreur)
            else:
                self.dateTime_new_date_loc.setStyleSheet(styleNormal)
                msg=MessageBoxValid()
                msg.setText("Information: \n\nLa date de location N°"+str(numero)+" est modifiée avec succée")
                msg.exec()
                self.saveLocation=False
                self.Data.update_date_loc(numero,datec)
                self.click_btn_date()
    
    #function btn_update_duree_loc_confirm
    def click_btn_update_duree_loc_confirm(self):
        numero='#'+self.line_numero_loc_date.text()
        duree=self.line_new_duree_loc.text()
        styleNormal="border-radius:10px;border: 2px solid rgb(133, 183, 198);background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        styleErreur="border-radius:10px;border: 2px solid red;background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        msgBox=MessageBoxErreur()
        if(numero=="" or numero not in self.Data.dictLocation.keys()):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Numéro\t\n*Champs Obligatoire\n*5 Chiffres\n*Pas de Location avec ce Numéro")
            msgBox.exec()
            self.line_numero_loc_date.setClearButtonEnabled(True)
            self.line_numero_loc_date.setStyleSheet(styleErreur)
            s1=False
        else:
            self.line_numero_loc_date.setClearButtonEnabled(False)
            self.line_numero_loc_date.setStyleSheet(styleNormal)
            s1=True
        if(duree=="" or int(duree)<0):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Durée\t\n*Champ Obligatoire\n*Durée Doit être supérieur à 1j ")
            msgBox.exec()
            self.line_new_duree_loc.setStyleSheet(styleErreur)
            self.line_new_duree_loc.setClearButtonEnabled(True)
            s2=False
        else:
            self.line_new_duree_loc.setStyleSheet(styleNormal)
            self.line_new_duree_loc.setClearButtonEnabled(False)
            s2=True
        if(s1 and s2):
            liste=[]
            datec=self.Data.dictLocation[numero].date
            s3=True
            car=self.Data.dictLocation[numero].matricule
            for key,item in self.Data.dictLocation.items():
                if(key!=numero and car==item.matricule):
                    liste.append(item)
            for item in liste:
                if(item.date<=datec<=(item.date).addDays(int(item.duree)) or datec<=item.date<=datec.addDays(int(duree))):
                    s3=False
                    prop=item
            if(s3==False):
                msgBox.setText("Erreur: \n\nLa voiture "+str(car)+" est déja louée pour cette date à Mr/Mrs "+str(self.Data.dictClient[prop.cin].nom)+" "+str(self.Data.dictClient[prop.cin].prenom))
                msgBox.exec()
                self.line_new_duree_loc.setStyleSheet(styleErreur)
                self.line_new_duree_loc.setClearButtonEnabled(True)
            else:
                self.line_new_duree_loc.setStyleSheet(styleNormal)
                self.line_new_duree_loc.setClearButtonEnabled(False)
                msg=MessageBoxValid()
                msg.setText("Information: \n\nLa date de location N°"+str(numero)+" est modifiée avec succée")
                msg.exec()
                self.saveLocation=False
                self.Data.update_duree_loc(numero,duree)
                self.click_btn_duree()

    #function btn_imprimer_loc_confirm
    def click_btn_imprimer_loc_confirm(self):
        numero="#"+str(self.line_numero_imprimer_loc.text())
        styleNormal="border-radius:10px;border: 2px solid rgb(133, 183, 198);background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        styleErreur="border-radius:10px;border: 2px solid red;background-color: rgb(46, 47, 65);color:white;margin-right: 0px;font: 57 14pt ""Oswald"";color: rgb(209, 208, 222);"
        msgBox=MessageBoxErreur()
        if(numero=="" or numero not in self.Data.dictLocation.keys()):
            msgBox.setText("Erreur de Saisie: \n\nVérifier le champ Numéro\t\n*Champs Obligatoire\n*5 Chiffres\n*Pas de Location avec ce Numéro")
            msgBox.exec()
            self.line_numero_loc_date.setClearButtonEnabled(True)
            self.line_numero_loc_date.setStyleSheet(styleErreur)
        else:
            self.line_numero_loc_date.setClearButtonEnabled(False)
            self.line_numero_loc_date.setStyleSheet(styleNormal)
            dirname = os.path.dirname(__file__)
            outputfile = os.path.join(dirname, "Facture\\"+str(numero)+".pdf")
            if(os.path.exists(outputfile)):
                subprocess.Popen(outputfile, shell=True)
            else:
                self.Data.imprimer_facture(numero)
            self.click_btn_imprimer()
    
    #function btn_imprimer
    def click_btn_imprimer(self):
        if(self.recupClient==False or self.recupLocation==False or self.recupVoiture==False):
            msgBox=MessageBoxErreur()
            msgBox.setText("Erreur: \n\nEssayer de récupérer les Données\n")
            msgBox.exec()
        else:
            self.line_numero_imprimer_loc.clear()
            self.Main_frame.setCurrentIndex(35)

    def click_maximize(self):
        if(self.isMaximized()==True):
            self.showNormal()
            self.tableWidget.horizontalHeader().setDefaultSectionSize(125)
            self.View_car.horizontalHeader().setDefaultSectionSize(135)
            self.View_car_marque.horizontalHeader().setDefaultSectionSize(135)
            self.View_car_color.horizontalHeader().setDefaultSectionSize(135)
            self.View_car_dispo.horizontalHeader().setDefaultSectionSize(135)
            self.View_car_louee.horizontalHeader().setDefaultSectionSize(135)
            self.View_car_loue.horizontalHeader().setDefaultSectionSize(135)
            self.View_Client.horizontalHeader().setDefaultSectionSize(135)
            self.View_location.horizontalHeader().setDefaultSectionSize(135)
            self.View_loc_cin.horizontalHeader().setDefaultSectionSize(135)
            self.View_loc_matricule.horizontalHeader().setDefaultSectionSize(135)
            self.View_loc_date.horizontalHeader().setDefaultSectionSize(135)
            self.View_car_dure.horizontalHeader().setDefaultSectionSize(135)
            self.View_search_2_date.horizontalHeader().setDefaultSectionSize(135)
        else:
            self.showMaximized()
            self.View_car_color.horizontalHeader().setDefaultSectionSize(230)
            self.tableWidget.horizontalHeader().setDefaultSectionSize(220)
            self.View_car.horizontalHeader().setDefaultSectionSize(230)
            self.View_car_marque.horizontalHeader().setDefaultSectionSize(230)
            self.View_car_dispo.horizontalHeader().setDefaultSectionSize(230)
            self.View_car_louee.horizontalHeader().setDefaultSectionSize(230)
            self.View_car_loue.horizontalHeader().setDefaultSectionSize(230)
            self.View_Client.horizontalHeader().setDefaultSectionSize(220)
            self.View_location.horizontalHeader().setDefaultSectionSize(220)
            self.View_loc_cin.horizontalHeader().setDefaultSectionSize(220)
            self.View_loc_matricule.horizontalHeader().setDefaultSectionSize(220)
            self.View_loc_date.horizontalHeader().setDefaultSectionSize(220)
            self.View_car_dure.horizontalHeader().setDefaultSectionSize(220)
            self.View_search_2_date.horizontalHeader().setDefaultSectionSize(220)
            self.View_loc_date.horizontalHeader().setDefaultSectionSize(220)     
    
    def click_exit(self):
        if(self.saveLocation==False or self.saveClient==False or self.saveVoiture==False):
            msgBox = QMessageBox()
            msgBox.setIcon(QMessageBox.Question)
            msgBox.setStandardButtons(QMessageBox.Yes | QMessageBox.No |QMessageBox.Cancel)
            msgBox.setWindowFlag(Qt.FramelessWindowHint)
            msgBox.setMinimumWidth(500)
            msgBox.setFixedWidth(600)
            msgBox.setWindowOpacity(0.85)
            msgBox.setText("Quitter: \n\nVoudriez-vous Enregistrer les Modifications avant de quitter ? \n")
            msgBox.setStyleSheet("QPushButton{border-radius:10px;border:2px solid rgb(99, 126, 193);background-color: rgb(85, 96, 124);font: 87 14pt ""Segoe UI"";width:60px;color: rgb(255, 255, 255);}QPushButton:hover{background-color: rgb(114, 153, 146);}QLabel{font: 87 14pt ""Segoe UI"";color: white;}QMessageBox{border-radius:25px;border:7px solid rgb(95,158,160);background-color: rgb(32, 33, 53);}")
            res=msgBox.exec()
            if(res==QMessageBox.Yes):
                self.Data.save_car()
                self.Data.save_client()
                self.Data.save_Location()
                msg=MessageBoxValid()
                msg.setText("Information: \n\n Modifications Enregistrer avec Sucées\n\n Au revoir...")
                msg.exec()
                exit()
            elif(res==QMessageBox.No):
                exit()
        else:
            exit()

    #function btn_accueil
    def click_btn_accueil(self):
        if(self.recupClient==False or self.recupLocation==False or self.recupVoiture==False):
            msgBox=MessageBoxErreur()
            msgBox.setText("Erreur: \n\nEssayer de récupérer les Données\n")
            msgBox.exec()
        else:
            self.Data.refresh_rent(self.label_118)
            self.label_79.setText(self.Data.benefice_dash())
            self.label_82.setText(str(len(self.Data.dictVoiture)))
            self.label_85.setText(str(len(self.Data.dictClient)))
            self.label_96.setText(str(self.Data.dispo_car()))
            self.label_99.setText(str(len(self.Data.dictVoiture)-self.Data.dispo_car()))
            self.label_93.setText(str(self.Data.most_car_rent()))
            self.label_92.setText(str(self.Data.most_client_rent()))
            self.label_94.setText(str(self.Data.revenu_today()))
            self.Data.list_cars_month(self.tableWidget)
            self.Main_frame.setCurrentIndex(31)

    #function btnFacebook
    def click_btnFacebook(self):
        webbrowser.open('https://www.facebook.com/haithem.benslama')

    #fucntion btnInstagram
    def click_btnInstagram(self):
        webbrowser.open('https://www.instagram.com/haithem_benslama/')

    #fucntion btnLinkedin
    def click_btnLinkedin(self):
        webbrowser.open('https://www.linkedin.com/in/haithem-ben-slama-89212822b/')

    #function btnRefrech
    def click_btnRefrech(self):
        if(self.recupClient==False or self.recupLocation==False or self.recupVoiture==False):
            msgBox=MessageBoxErreur()
            msgBox.setText("Erreur: \n\nEssayer de récupérer les Données\n")
            msgBox.exec()
        else:
            if(self.Main_frame.currentIndex()==31):
                self.click_btn_accueil()
            else:
                self.Data.refresh_rent(self.label_118)
    
    def __init__(self):
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.maximize.clicked.connect(self.click_maximize)
        self.hide.clicked.connect(self.showMinimized)
        self.close.clicked.connect(self.click_exit)
        self.setWindowFlag(Qt.FramelessWindowHint)
        regexAlpha=QtCore.QRegExp("[a-z A-Z]+")
        validatorAlpha =QtGui.QRegExpValidator(regexAlpha)
        #Validator Champs Voiture
        self.line_prix.setValidator(QIntValidator())
        self.line_new_prix.setValidator(QIntValidator())
        #Validator Champs Client
        self.line_cin_adresse.setValidator(QIntValidator())
        self.line_cin_tel.setValidator(QIntValidator())
        self.line_new_tel.setValidator(QIntValidator())
        self.line_cin_mail.setValidator(QIntValidator())
        self.line_nom_add.setValidator(validatorAlpha)
        self.line_cin_search.setValidator(QIntValidator())
        self.line_Cin_add.setValidator(QIntValidator())
        self.line_cin_delete.setValidator(QIntValidator())
        self.line_prenom_add.setValidator(validatorAlpha)
        self.line_age_add.setValidator(QIntValidator())
        self.line_numero_add.setValidator(QIntValidator())
        self.line_cin_delete.setValidator(QIntValidator())
        self.line_cin_adresse.setValidator(QIntValidator())
        self.line_cin_tel.setValidator(QIntValidator())
        self.line_new_tel.setValidator(QIntValidator())
        self.line_cin_mail.setValidator(QIntValidator())
        self.comboBox_add_location.currentTextChanged.connect(lambda:self.Data.calcul_montant(self.comboBox_add_location,self.line_dure_add_loc,self.line_montant_loc))
        #Validator Champs location
        self.line_cin_add_loc.setValidator(QIntValidator())
        self.line_dure_add_loc.setValidator(QIntValidator())
        self.line_numero_del_loc.setValidator(QIntValidator())
        self.line_numero_update_loc.setValidator(QIntValidator())
        self.line_numero_loc_date.setValidator(QIntValidator())
        self.line_new_duree_loc.setValidator(QIntValidator())
        self.line_cin_search_loc.setValidator(QIntValidator())
        self.line_duree_search_loc.setValidator(QIntValidator())
        self.line_Cin_add_2.setValidator(QIntValidator())
        self.line_age_add_2.setValidator(QIntValidator())
        self.line_numero_add_2.setValidator(QIntValidator())
        self.line_nom_add_2.setValidator(validatorAlpha)
        self.line_prenom_add_2.setValidator(validatorAlpha)
        self.line_numero_imprimer_loc.setValidator(QIntValidator())
        #btn_add_client
        self.btn_add_client.clicked.connect(self.click_btn_add_client)

        #btn_delete_client
        self.btn_delete_client.clicked.connect(self.click_btn_delete_client)

        #btn_update_client
        self.btn_update_client.clicked.connect(self.click_btn_update_client)

        #btn_adresse
        self.btn_adresse.clicked.connect(self.click_btn_adresse)
        
        #btn_tel
        self.btn_tel.clicked.connect(self.click_btn_tel)

        #btn_mail
        self.btn_mail.clicked.connect(self.click_btn_mail)

        #btn_affiche_client
        self.btn_affiche_client.clicked.connect(self.click_btn_affiche_client)

        #btn_search_client
        self.btn_search_client.clicked.connect(self.click_btn_search_client)

        #btn_add_voiture
        self.btn_add_voiture.clicked.connect(self.click_btn_add_voiture)

        #btn_delete_voiture
        self.btn_delete_voiture.clicked.connect(self.click_btn_delete_voiture)

        #btn_matricule
        self.btn_matricule.clicked.connect(self.click_btn_matricule)

        #btn_marque
        self.btn_marque.clicked.connect(self.click_btn_marque)

        #btn_age
        self.btn_age.clicked.connect(self.click_btn_age)

        #btn_update_voiture
        self.btn_update_voiture.clicked.connect(self.click_btn_update_voiture)

        #btn_prix
        self.btn_prix.clicked.connect(self.click_btn_prix)

        #btn_couleur
        self.btn_couleur.clicked.connect(self.click_btn_couleur)

        #btn_search_voiture
        self.btn_search_voiture.clicked.connect(self.click_btn_search_voiture)

        #btn_matricule_voiture
        self.btn_matricule_voiture.clicked.connect(self.click_btn_matricule_voiture)

        #btn_marque_voiture
        self.btn_marque_voiture.clicked.connect(self.click_btn_marque_voiture)

        #btn_couleur_voiture
        self.btn_couleur_voiture.clicked.connect(self.click_btn_couleur_voiture)

        #btn_disponible_voiture
        self.btn_disponible_voiture.clicked.connect(self.click_btn_disponible_voiture)

        #btn_louee_voiture
        self.btn_louee_voiture.clicked.connect(self.click_btn_louee_voiture)

        #btn_louee_2_voiture
        self.btn_louee_2_voiture.clicked.connect(self.click_btn_louee_2_voiture)

        #btn_add_location
        self.btn_add_location.clicked.connect(self.click_btn_add_location)
        
        #btn_delete_location
        self.btn_delete_location.clicked.connect(self.click_btn_delete_location)
        
        #btn_update_location
        self.btn_update_location.clicked.connect(self.click_btn_update_location)

        #btn_date
        self.btn_date.clicked.connect(self.click_btn_date)

        #btn_duree
        self.btn_duree.clicked.connect(self.click_btn_duree)

        #btn_afficher_location
        self.btn_afficher_location.clicked.connect(self.click_btn_afficher_location)

        #btn_recherche_location
        self.btn_recherche_location.clicked.connect(self.click_btn_recherche_location)

        #btn_cin_loc
        self.btn_cin_loc.clicked.connect(self.click_btn_recherche_location)

        #btn_matricule_loc
        self.btn_matricule_loc.clicked.connect(self.click_btn_matricule_loc)

        #btn_date_loc
        self.btn_date_loc.clicked.connect(self.click_btn_date_loc)

        #btn_duree_loc
        self.btn_duree_loc.clicked.connect(self.click_btn_duree_loc)

        #btn_2_date
        self.btn_2_date.clicked.connect(self.click_btn_2_date)

        #btn_back_location
        self.btn_back_location.clicked.connect(self.click_back)

        #btn_back_client
        self.btn_back_client.clicked.connect(self.click_back)
        
        #btn_back_voiture
        self.btn_back_voiture.clicked.connect(self.click_back)

        #btn_client
        self.btn_client.clicked.connect(self.click_btn_client)

        #btn_voiture
        self.btn_voiture.clicked.connect(self.click_btn_voiture)

        #btn_location
        self.btn_location.clicked.connect(self.click_btn_location)

        #btn_recuperer
        self.btn_recuperer.clicked.connect(lambda:self.Main_frame.setCurrentIndex(32))

        #btn_load_client
        self.btn_load_client.clicked.connect(self.click_btn_load_client)

        #btn_load_car
        self.btn_load_car.clicked.connect(self.click_btn_load_car)

        #btn_load_location
        self.btn_load_location.clicked.connect(self.click_btn_load_location)

        #btn_load_all
        self.btn_load_all.clicked.connect(self.click_btn_load_all)

        #btn_sauvgarder
        self.btn_sauvgarder.clicked.connect(lambda:self.Main_frame.setCurrentIndex(33))

        #btn_save_client
        self.btn_save_client.clicked.connect(self.click_btn_save_client)

        #btn_save_car
        self.btn_save_car.clicked.connect(self.click_btn_save_car)

        #btn_save_location
        self.btn_save_location.clicked.connect(self.click_btn_save_location)

        #btn_save_all
        self.btn_save_all.clicked.connect(self.click_btn_save_all)

        #btn_add_car
        self.btn_add_car.clicked.connect(self.click_btn_add_car)

        #btn_delete_car
        self.btn_delete_car.clicked.connect(self.click_btn_delete_car)

        #btn_list_car
        self.btn_list_car.clicked.connect(self.click_btn_list_car)

        #btn_delete_car_confirm
        self.btn_delete_car_confirm.clicked.connect(self.click_btn_delete_car_confirm)

        #btn_delete_age_confirm
        self.btn_delete_age_confirm.clicked.connect(self.click_btn_delete_age_confirm)

        #btn_update_price_confirm
        self.btn_update_price_confirm.clicked.connect(self.click_btn_update_price_confirm)
        
        #click_btn_update_color_confirm
        self.btn_update_color_confirm.clicked.connect(self.click_btn_update_color_confirm)

        #btn_search_car_matricule_confirm
        self.btn_search_car_matricule_confirm.clicked.connect(self.click_btn_search_car_matricule_confirm)

        #btn_search_marque_confirm
        self.btn_search_marque_confirm.clicked.connect(self.click_btn_search_marque_confirm)

        #btn_search_color_confirm
        self.btn_search_color_confirm.clicked.connect(self.click_btn_search_color_confirm)

        #btn_add_client_confirm
        self.btn_add_client_confirm.clicked.connect(self.click_btn_add_client_confirm)

        #btn_delete_client_confirm
        self.btn_delete_client_confirm.clicked.connect(self.click_btn_delete_client_confirm)

        #btn_update_adresse
        self.btn_update_adresse.clicked.connect(self.click_btn_update_adresse)

        #btn_update_tel
        self.btn_update_tel.clicked.connect(self.click_btn_update_tel)

        #btn_update_mail
        self.btn_update_mail.clicked.connect(self.click_btn_update_mail)

        #btn_search_with_cin
        self.btn_search_with_cin.clicked.connect(self.click_btn_search_with_cin)

        #btn_add_location_next
        self.btn_add_location_next.clicked.connect(self.click_btn_add_location_next)
        #btn_confirm_add_location
        self.btn_confirm_add_location.clicked.connect(self.click_btn_confirm_add_location)

        #btn_add_client_location
        self.btn_add_client_location.clicked.connect(self.click_btn_add_client_location)

        #btn_add_client_confirm__from_loc
        self.btn_add_client_confirm__from_loc.clicked.connect(self.click_btn_add_client_confirm__from_loc)
        #btn_delete_loc_confirm
        self.btn_delete_loc_confirm.clicked.connect(self.click_btn_delete_loc_confirm)

        #btn_search_loc_cin_confirm
        self.btn_search_loc_cin_confirm.clicked.connect(self.click_btn_search_loc_cin_confirm)

        #btn_search_matricule_loc_confirm
        self.btn_search_matricule_loc_confirm.clicked.connect(self.click_btn_search_matricule_loc_confirm)

        #btn_serach_date_location
        self.btn_serach_date_location.clicked.connect(self.click_btn_serach_date_location)

        #btn_search_loc_dure_confirm
        self.btn_search_loc_dure_confirm.clicked.connect(self.click_btn_search_loc_dure_confirm)
        
        #btn_search_2_date_confirm
        self.btn_search_2_date_confirm.clicked.connect(self.click_btn_search_2_date_confirm)

        #btn_search_car_louee
        self.btn_search_car_louee.clicked.connect(self.click_btn_search_car_louee)

        #btn_update_date_loc_confirm
        self.btn_update_date_loc_confirm.clicked.connect(self.click_btn_update_date_loc_confirm)

        #btn_update_duree_loc_confirm
        self.btn_update_duree_loc_confirm.clicked.connect(self.click_btn_update_duree_loc_confirm)

        #btn_imprimer_loc_confirm
        self.btn_imprimer_loc_confirm.clicked.connect(self.click_btn_imprimer_loc_confirm)

        #btn_imprimer
        self.btn_imprimer.clicked.connect(self.click_btn_imprimer)

        #btn_accueil
        self.btn_accueil.clicked.connect(self.click_btn_accueil)

        #btnFacebook
        self.btnFacebook.clicked.connect(self.click_btnFacebook)

        #btnInstagram
        self.btnInstagram.clicked.connect(self.click_btnInstagram)

        #btnLinkedin
        self.btnLinkedin.clicked.connect(self.click_btnLinkedin)

        #btnRefrech
        self.btnRefrech.clicked.connect(self.click_btnRefrech)
        #btn_about
        self.btn_about.clicked.connect(lambda:self.Main_frame.setCurrentIndex(37))

        #btn_recup
        self.btn_recup.clicked.connect(lambda:self.Main_frame.setCurrentIndex(32))

        #btn_back_location_1
        self.btn_back_location_1.clicked.connect(lambda:self.Main_frame.setCurrentIndex(20))

        #btn_exit
        self.btn_exit.clicked.connect(self.click_exit)
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = LoadingScreen()
    sys.exit(app.exec_())