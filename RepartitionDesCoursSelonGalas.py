import AssoConnectProcess
# ce script sert à lister les cours
planning_semaine = AssoConnectProcess.fill_planning()

print(type(planning_semaine))
print(planning_semaine)

print("[cours]")

for jour in planning_semaine:
    for i in jour:
        nb_garcon = 0
        for e in i.liste_eleves:
            if e["Genre"] == "Masculin":
                nb_garcon+=1
        print(f"[cours.{i.jour} {i.heure} {i.prof['diminutif']} {i.discipline} {i.niveau}]")
        print(f"    jour = '{i.jour}'")
        print(f"    heure = '{i.heure}'")
        print(f"    prof = '{i.prof['nom']}'")
        print(f"    discipline = '{i.discipline}'")
        print(f"    niveau = '{i.niveau}'")
        print(f"    nombre_eleves = '{i.nb_eleves}'")
        print(f"    gala = ")
        print(f"    loge = ''")
        print(f"    ordre_gala = ")
        print(f"    changement_de_costumes = ''")
        print(f"    nb_garcons = {str(nb_garcon)}")
        print(f"    couleur_sac = ''")
