import toml
import pandas as pd
import selenium.common.exceptions
from selenium import webdriver
from selenium.webdriver.common.by import By
import os
from time import sleep
import unidecode

from GetDatabase import get_assoconnect_data_base

liste_profs = toml.load("parameters.toml")["profs"]
liste_couleurs = toml.load("parameters.toml")["couleurs"]


class Eleve:
    def __init__(self, nom: str, prenom: str, id: int, token: int):
        """
        Chaque élève est marqué par son nom, son prénom et son ID unique. Le Token permet de définir lequel des deux
        doublons est le plus pertinent de garder pour servir de base pour la fusion des cours
        :param nom:
        :param prenom:
        :param id:
        :param token:
        """
        self.nom = nom.upper()
        self.prenom = prenom.capitalize()
        self.id = id
        self.token = token

    # redéfinition des comparateurs >, < & == pour comparer directement les tokens.
    def __eq__(self, other):
        if self.token == other.token:
            return True
        else:
            return False

    def __gt__(self, other):
        if self.token > other.token:
            return True
        else:
            return False

    def __lt__(self, other):
        if self.token < other.token:
            return True
        else:
            return False

    def __repr__(self):
        return f"{self.nom} {self.prenom} {str(self.token)}, {str(self.id)}"



# RÉCUPÉRATION DE LA BASE DE DONNÉE SUR ASSOCONNECT
#workbook = get_assoconnect_data_base()
workbook = pd.read_excel("export.xlsx")
# liste des informations à garder comme étant relevantes et pertinentes pour définir le compte principal
tokens = ["ID du Contact", "Nom", "Prénom", "Statut adhérent", "Statut donateur", "Email", "Date de naissance",
          "Téléphone fixe", "Téléphone mobile", "Sexe", "Adresse", "Code postal", "Ville",
          "Email mère ou responsable légal 1", "Email père ou responsable légal 2", "Tél responsable légal 1",
          "Tél responsable légal 2", "Urgence Nom personne à prévenir", "Urgence Tél à composer",
          "Urgence Autorisation hospitalisation",
          "Indications médicales RAS ou allergies, asthme, dificultés auditives, difficultés visuelles ..."]

# création de la liste d'élèves
liste_eleve = []
for index in workbook.index:
    nom = unidecode.unidecode(str(workbook.loc[index, "Nom"]))
    prenom = unidecode.unidecode(str(workbook.loc[index, "Prénom"]))
    id = workbook.loc[index, "ID du Contact"]
    token = 0
    for i in tokens:
        if pd.notna(workbook.loc[index, i]):
            token += 1

    liste_eleve.append(Eleve(nom, prenom, id, token))

# on fait une liste avec seulement les noms et prénom pour trouver les doublons
liste_eleve = sorted(liste_eleve, key=lambda x: x.nom)


# on créé la liste des doublons, ceux qui ont le même nom & le même prénom mais pas le même id
seen_twice = []

for i in liste_eleve:
    liste_eleve.remove(i)
    for e in liste_eleve:
        if i.nom == e.nom and i.prenom == e.prenom and i.id != e.id:
            seen_twice.append([i, e])
for i in seen_twice:
    print(i)
print(f"Il y a {len(seen_twice)} doublons trouvés dans AssoConnect")


def unlock_pass(password: str) -> str:
    charac = "2345689!=%°+"
    password = password[::-1]
    for c in charac:
        password = password.replace(c, "")
    return password

# connection à AssoConnect
params = toml.load("./parameters.toml")
password = unlock_pass(params["AssoConnect"]["password"])
mail = params["AssoConnect"]["mail"]
url = params["AssoConnect"]["connection_page"]
current_dir = os.getcwd()

chrome_options = webdriver.ChromeOptions()
# chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument("--disable-gpu")
prefs = {'download.default_directory': current_dir}
chrome_options.add_experimental_option('prefs', prefs)

driver = webdriver.Chrome(chrome_options=chrome_options, options=chrome_options)
# direction vers la page
driver.set_window_size(1900, 1060)
driver.get(url)

# entrée des codes & validation
driver.find_element(By.XPATH, """/html/body/div/div[3]/div[1]/div[1]/form/div[1]/input""").send_keys(mail)
driver.find_element(By.XPATH, """/html/body/div/div[3]/div[1]/div[1]/form/button/span""").click()

driver.find_element(By.XPATH, """/html/body/div/div[3]/div[1]/div[1]/form/div[3]/input""").send_keys(password)
driver.find_element(By.XPATH, """/html/body/div/div[3]/div[1]/div[1]/form/button/span""").click()


# FUSION DES COMPTES
tot = len(seen_twice)
trot = 0
for i in seen_twice:
    trot += 1
    # entrée sur la page de fusion
    driver.get("https://academie-de-ballet-nini-theilade.assoconnect.com/platform/contacts/advanced/272677")
    if i[0] > i[1]:
        premier = i[0]
        second = i[1]
    else:
        premier = i[1]
        second = i[0]
    # sélection du champ 1
    driver.find_element(By.XPATH, '//*[@id="BuyPackerUser1IdHelper"]').send_keys(str(premier.id))
    sleep(2)
    for i in range (1, 200):
        try:
            try:
                driver.find_element(By.XPATH, f"""//*[@id="ui-id-{str(i)}"]""").click()
            except selenium.common.exceptions.ElementNotInteractableException:
                pass
        except selenium.common.exceptions.NoSuchElementException:
            pass
    driver.find_element(By.XPATH, '//*[@id="BuyPackerUser2IdHelper"]').send_keys(str(second.id))
    sleep(2)
    for i in range (1, 200):
        try:
            try:
                driver.find_element(By.XPATH, f"""//*[@id="ui-id-{str(i)}"]""").click()
            except selenium.common.exceptions.ElementNotInteractableException:
                pass
        except selenium.common.exceptions.NoSuchElementException:
            pass
    driver.find_element(By.XPATH, """//*[@id="fusionSubmit"]""").click()
    print(f"({str(trot)}/{str(tot)})le compte {str(premier.id)} de {premier.nom} {premier.prenom} a été fusionné avec le compte {str(second.id)} ")
print("Fusion terminée !")
driver.close()