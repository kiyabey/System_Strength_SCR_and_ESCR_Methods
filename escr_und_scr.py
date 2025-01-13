from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from math import sqrt

# Zahlen mit SI-Präfixen
K = 10**3        # 1 Kilo = 1.000
M = 10**6        # 1 Mega = 1.000.000
G = 10**9        # 1 Giga = 1.000.000.000
T = 10**12       # 1 Tera = 1.000.000.000.000

# Lade die Excel-Dateien
wb_knoten = load_workbook('Ergebnisknoten_3phKurzschluss_Kopie.xlsx')
wb_komponenten = load_workbook('Komponenten_Kopie.xlsx')

# Wähle die Arbeitsblätter aus
ws_knoten = wb_knoten['Ergebnisknoten_3phKurzschluss']
ws_komponenten = wb_komponenten['Komponenten']


#ESCR (Equivalent Short Circuit Ratio) Methode


def escr(ws_1, ws_2):
    
    """
    Funktion:
    Diese Funktion berechnet die Kurzschlussleistungsverhältnis nach der ESCR Methode und speichert die ESCR-Werte in der Excel-Datei
    
    Parameter:
    ws_1: Das Arbeitsblatt, in dem sich die benötigten Informationen (ausser Nennleistung) befinden und die ESCR-werte gespeichert werden (Knoten)
    ws_2: Das Arbeitsblatt, in dem sich die Nennleistung der Stromrichter befinden (Komponenten)
    """

  

    #Neue Spalte für ESCR-Werte
    neue_spalte = ws_1.max_column + 1
    ws_1.insert_cols(neue_spalte)
    spalte_buchstabe = get_column_letter(neue_spalte)
    ws_1[f'{spalte_buchstabe}' + '8'].value = 'ESCR'
    
    
    """
    Die Interaktionsfaktoren stehen noch nicht zur Verfügung.
    """

    #Interaktionsfaktor
    IF = 1
   
    #ESCR-Wert berechnen
     
    # Iteration durch die Zeilen in der Eingabespalte
    
    for zeile in range(10, ws_1.max_row + 1):
       summe = 0     #Summe der Multipikationen von Iterationsfaktoren und Nennleistung der Stromrichter für alle Knoten
       p_n_sr_i = 0  #Nennleistung der Stromrichter am Knoten, für den man den ESCR-Wert berechnet


       u_n = float(ws_1['e' + str(zeile)].value) * K  #Nennspannung
      
       # Überprüfen, ob der Wert(Nennleistung) in der Zelle leer ist
       if u_n is None:
         break

       i_k = float(ws_1['k' + str(zeile)].value) * K / 10**5  #Anfangs-Kurzschlusswechselstrom                         
       s_k = sqrt(3)*u_n*i_k                                  #Subtransiente Kurzschlussleistung 
       
       #Initialisiere p_n_sr_i (Anhand des Vergleichs der Knotennamen in Excel-Dateien)
       for zeile_4 in range(4, ws_2.max_row + 1):
            if ws_2['dd' + str(zeile_4)].value is None:
               break
            
            if ws_1['a' + str(zeile)].value == ws_2['dd' + str(zeile_4)].value :
               p_n_sr_i = float(ws_2['ae' + str(zeile_4)].value) * M
               break
       

       # Iteraktionsfaktor * Nennleistung der Stromrichter 
       for zeile_2 in range(10, ws_1.max_row + 1):
          p_n_sr_j = 0  #Nennleistung der Stromrichter an einem anderen Knoten 
         
          # Überprüfe, ob die leere Zelle erreicht wird (Anhand der Spannungsspalte)
          u_n_2 = ws_1['e' + str(zeile_2)].value   #Nennspannung
         
          if u_n_2 is None:
             break
          
          # Überprüfe, ob die eigene Zeile des Knotens erreich wird
          if zeile_2 == zeile:
             continue
         
          #Initialisiere p_n_sr_j (Anhand des Vergleichs der Knotennamen in Excel-Dateien)
          for zeile_3 in range(4, ws_2.max_row + 1):
             if ws_2['dd' + str(zeile_3)].value is None:
                break
            
             if ws_1['a' + str(zeile_2)].value == ws_2['dd' + str(zeile_3)].value :
                p_n_sr_j = float(ws_2['ae' + str(zeile_3)].value) * M
                break
         
          #Iteration durchführen 
          summe += p_n_sr_j * IF
 
       
       if (p_n_sr_i + summe) == 0:
          escr_wert = None
          print('Division durch Null!' + str(zeile))
       else:
          escr_wert = s_k / (p_n_sr_i + summe)
          
       
       
       #ESCR-Werte in der neuen Spalte speichern
       ws_1[f'{spalte_buchstabe}' + str(zeile)].value = escr_wert
       

#SCR (Short Circuit Ratio) Methode
def scr(ws_1, ws_2):
    
    """
    Funktion:
    Diese Funktion berechnet die Kurzschlussleistungsverhältnis nach der SCR Methode und speichert die SCR-Werte in der Excel-Datei
    
    Parameter:
    ws_1: Das Arbeitsblatt, in dem sich die benötigten Informationen (ausser Nennleistung) befinden und die ESCR-werte gespeichert werden (Knoten)
    ws_2: Das Arbeitsblatt, in dem sich die Nennleistung der Stromrichter befinden (Komponenten)
    """

  

    #Neue Spalte für ESCR-Werte
    neue_spalte = ws_1.max_column + 1
    ws_1.insert_cols(neue_spalte)
    spalte_buchstabe = get_column_letter(neue_spalte)
    ws_1[f'{spalte_buchstabe}' + '8'].value = 'SCR'
    
    
    #SCR-Wert berechnen
     
    # Iteration durch die Zeilen in der Eingabespalte
    
    for zeile in range(10, ws_1.max_row + 1):
       p_n_sr_i = 0  #Nennleistung der Stromrichter am Knoten, für den man den ESCR-Wert berechnet


       u_n = float(ws_1['e' + str(zeile)].value) * K  #Nennspannung
      
       # Überprüfen, ob der Wert(Nennleistung) in der Zelle leer ist
       if u_n is None:
         break

       i_k = float(ws_1['k' + str(zeile)].value) * K / 10**5  #Anfangs-Kurzschlusswechselstrom                         
       s_k = sqrt(3)*u_n*i_k                                  #Subtransiente Kurzschlussleistung 
       
       #Initialisiere p_n_sr_i (Anhand des Vergleichs der Knotennamen in Excel-Dateien)
       for zeile_4 in range(4, ws_2.max_row + 1):
            if ws_2['dd' + str(zeile_4)].value is None:
               break
            
            if ws_1['a' + str(zeile)].value == ws_2['dd' + str(zeile_4)].value :
               p_n_sr_i = float(ws_2['ae' + str(zeile_4)].value) * M
               break
 
       
       if p_n_sr_i == 0:
          scr_wert = None
          print('Division durch Null!' + str(zeile))
       else:
          scr_wert = s_k / (p_n_sr_i)
          
       
       
       #SCR-Werte in der neuen Spalte speichern
       ws_1[f'{spalte_buchstabe}' + str(zeile)].value = scr_wert
       



  
#escr(ws_knoten, ws_komponenten)
#scr(ws_knoten, ws_komponenten)
wb_knoten.save('Ergebnisknoten_3phKurzschluss_Kopie.xlsx')


