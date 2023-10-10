import AssoConnectProcess
import graphviz
import itertools


planning_semaine = AssoConnectProcess.var_semaine




class EnfantPlusieursCours:
    def __init__(self, id:int, name:str, family_name:str, cours:list):
        """
        :param id: id Assoconnect de l'enfant
        :param name: prénom de l'enfant
        :param family_name: nom de famille de l'enfant
        :param cours: liste de cours dans lequel se trouve l'enfant
        """
        self.id = id
        self.name = name
        self.family_name = family_name
        self.cours = cours


    def __repr__(self):
        impr = (str(self.id) + f" a {len(self.cours)} cours")
        return impr

# On recherche tous les enfants qui sont dans plusieurs cours et on ajoute tous les cours qu'ils font
# dans leur liste de cours respectives (objet EnfantPlusieursCours.cours)
links = []
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
                                                      family_name=eleve["Nom"], cours=[eleve["cours"]]))
                else:
                    find = False
                    for e in links:
                        # si l'élève est déjà listé alors on complète
                        # ses cours dans sa liste (EnfantPlusieursCours.cours)
                        if e.id == eleve["ID"]:
                            e.cours.append(eleve["cours"])
                            find = True
                            break
                    # si l'élève n'est pas déjà listé alors on l'ajoute dans la liste (links)
                    if find is False:
                        links.append(EnfantPlusieursCours(id=eleve["ID"], name=eleve["Prénom"],
                                                          family_name=eleve["Nom"], cours=[eleve["cours"]]))

print(f"Tous les enfants qui ont plusieurs cours ont été trouvés. \n{str(len(links))} enfants ont plusieurs cours")

# on tri la liste des enfants par leur id
links.sort(key=lambda x: x.id)

# instanciation du graph
dot = graphviz.Graph(comment='Enfants qui font plusieurs cours')
dot.graph_attr["size"] = "10.25,7.75!"
# on liste tous les liens et on dénombre leur fréquence d'apparition
total = []
for l in links:
    combinations = list(itertools.combinations(l.cours, 2))
    for i in combinations:
        found = False
        # si le lien est déjà instancié dans la liste on augmente sa fréquence d'apparition (count)
        for e in total:
            if i == e["link"]:
                e["count"] += 1
                e["eleves"].append(f"{l.name} {l.family_name}")
                found = True
        # si le lien n'est pas déjà instancié alors on l'ajoute dans la liste
        if not found:
            total.append({"link":i, "count":1, "eleves":[f"{l.name} {l.family_name}"]})

print(f"Tous les liens entre les cours ont été trouvés.\nIl y a {str(len(total))} liens entre les cours à prendre en compte")

# On ajoute tous les cours au graph

for i in links:
    for e in i.cours:
        dot.node(f"{e.jour} {e.heure} {e.prof['diminutif']} {e.discipline}", style='filled',
                 fillcolor="#"+AssoConnectProcess.liste_couleurs[e.prof["fond"]],
                 fontcolor="#"+AssoConnectProcess.liste_couleurs[e.prof["couleur"]])

# on ajoute tous les liens au graph
for lien in total:
    eleves = lien["eleves"]
    lien = list(lien["link"])
    first = lien[0]
    firststr = f"{first.jour} {first.heure} {first.prof['diminutif']} {first.discipline}"
    second = lien[1]
    secondstr = f"{second.jour} {second.heure} {second.prof['diminutif']} {second.discipline}"
    if first.jour == second.jour:
        dot.edge(firststr, secondstr, label="\n".join([str(item) for item in eleves]), color="red", weight=str(float(len(eleves))))
    else:
        dot.edge(firststr, secondstr, label="\n".join([str(item) for item in eleves]))
print(dot)
dot.render('graph.pdf')
