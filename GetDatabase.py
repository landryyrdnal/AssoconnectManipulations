from time import sleep

import selenium.common.exceptions
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import os
import toml


def get_assoconnect_data_base()-> pd.DataFrame:
    """
    Cette fonction va chercher sur AssoConnect la dernière mise à jour de toutes les données et retourne un dataframe.
    :return: un dataframe Pandas avec toutes les données d'AssoConnect actualisées
    """

    def unlock_pass(password: str) -> str:
        charac = "2345689!=%°+"
        password = password[::-1]
        for c in charac:
            password = password.replace(c, "")
        return password

    file = os.getcwd()+r"/export.xlsx"
    params = toml.load("./parameters.toml")
    password = unlock_pass(params["AssoConnect"]["password"])
    mail = params["AssoConnect"]["mail"]
    url = params["AssoConnect"]["connection_page"]
    current_dir = os.getcwd()

    chrome_options = webdriver.ChromeOptions()
    #chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-gpu")
    prefs = {'download.default_directory' : current_dir}
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
    print("connection OK")

    # cliquer sur administration
    driver.find_element(By.XPATH, """/html/body/div[1]/div[3]/div[2]/a/div/span""").click()
    print("entrée sur la page ADMINISTRATION OK")

    # cliquer sur communauté

    try:
        driver.find_element(By.XPATH, """/html/body/div/div[3]/section/div/div[3]/div/div[4]/div[5]/a/div""").click()
    except selenium.common.exceptions.WebDriverException:
        sleep(4)
    print("entrée sur la page CONTACTS OK")

    # cliquer sur le boutton checker
    sleep(1)
    driver.find_element(By.XPATH,"""/html/body/div[1]/div[3]/section/div/div[4]/div/div[7]/div[4]/div/div[1]/div/div[2]/label/span[1]/span[2]""").click()
    print("checker tous les éléments OK")

    # tout sélectionner
    sleep(3)
    driver.find_element(By.XPATH,"""/html/body/div[1]/div[3]/section/div/div[4]/div/div[7]/span/div/u""").click()
    print("Sélection de tous les éléments OK")

    # enregistrer sous XLSX
    driver.find_element(By.XPATH, """/html/body/div[1]/div[3]/section/div/div[4]/div/div[7]/div[2]""").click()
    driver.find_element(By.XPATH, """/html/body/div[13]/div/ul/ul/li[1]/div/span""").click()
    print("Choix de l'export OK")
    sleep(3)
    # on supprime l'ancienne base de donnée .xlsx si elle existe
    try:
        os.remove(file)
        print("ancien fichier export.xlsx supprimé")
    except FileNotFoundError:
        print("aucun fichier export.xlsx à supprimer")
    # radio boutton pour sélectionner toutes les colonnes
    while not driver.find_element(By.XPATH, """/html/body/div[23]/div[1]/form/div/div[2]/div[1]/label/span[2]"""):
        sleep(1)
    driver.find_element(By.XPATH, """/html/body/div[23]/div[1]/form/div/div[2]/div[1]/label/span[2]""").click()
    print("Sélection de toutes les colonnes OK")

    # continuer
    driver.find_element(By.XPATH, """/html/body/div[23]/div[2]/div/div[2]""").click()
    print("validation OK")
    # on attends tant que le fichier n'est pas téléchargé
    while not os.path.isfile(file):
        sleep(1)
    print("téléchargement terminé")
    driver.close()
    database = pd.read_excel(file)
    return database




