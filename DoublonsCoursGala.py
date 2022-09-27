import AssoConnectProcess
import pandas as pd
import graphviz
import re

# workbook = pd.dataframe
#workbook = AssoConnectProcess.get_assoconnect_data_base()

#workbook = pd.read_excel("export.xlsx")
liste_niveaux = AssoConnectProcess.liste_niveaux
planning_semaine = AssoConnectProcess.var_semaine
links = []

class EnfantPlusieursCours:
    def __init__(self, id:int, name:str, family_name:str, cours:list):
        self.id = id
        self.name = name
        self.family_name = family_name
        self.cours = cours

    def __repr__(self):
        impr = (str(self.id) + f" a {len(self.cours)} cours")
        return impr




for jour in planning_semaine:
    for cours in jour:
        for eleve in cours.liste_eleves:
            if eleve['Autres cours'] != "":
                autres_cours = eleve["Autres cours"].split("|")
                for i in autres_cours:
                    if i == "" or i == ", " or i==" ":
                        autres_cours.remove(i)
                if len(links) == 0:
                    links.append(EnfantPlusieursCours(id=eleve["ID"],name=eleve["Prénom"],
                                                      family_name=eleve["Nom"], cours=autres_cours))
                else:
                    find = False
                    for e in links:
                        if e.id == eleve["ID"]:
                            print(f"{e.id} est complété")
                            for a in autres_cours:
                                e.cours.append(a)
                            find = True
                            break
                    if find is False:
                        print(f"{str(e.id)} est ajouté pour la première fois")
                        links.append(EnfantPlusieursCours(id=eleve["ID"], name=eleve["Prénom"],
                                                          family_name=eleve["Nom"], cours=autres_cours))


links.sort(key=lambda x: x.id)

for i in links:
    print(i.cours)

dot = graphviz.Digraph(comment='Enfants qui font plusieurs cours')
