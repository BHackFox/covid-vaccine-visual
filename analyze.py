import json         #scrittura e lettura di dizionari
import requests     #libreria per request di tipo GET/POST/PUT
import xlsxwriter   #creazione di file xlsx


def importer():

    getGithubUserContent()  #si scarica il file somministrazioni-vaccini-summary-latest.json e lo si salva

    with open("somministrazioni-vaccini-summary-latest.json") as js:
        jsondata = json.load(js)

    data = {}   #creazione di un dizionario vuoto
    for regione in jsondata:                                                                        #iterazione su ogni regione presente nel file .json
        data[jsondata[regione][0]["area"]] = {}                                                     #nome della regione
        data[jsondata[regione][0]["area"]]["nome_area"] = jsondata[regione][0]["nome_area"]         #nome identificativo == regione
        data[jsondata[regione][0]["area"]]["data"] = []                                             #creazione di una lista per archiviare le date di somministrazione
        data[jsondata[regione][0]["area"]]["prima_dose"] = []                                       #stessa cosa ma per le prime dosi
        data[jsondata[regione][0]["area"]]["seconda_dose"] = []                                     #per le seconde dosi
        data[jsondata[regione][0]["area"]]["booster"] = []                                          #per le dosi booster

        reg_data_somministrazione = data[jsondata[regione][0]["area"]]["data"]                      #utilizzo di abbreviazioni per rendere il codice più comprensibile
        reg_prima_dose = data[jsondata[regione][0]["area"]]["prima_dose"]
        reg_seconda_dose = data[jsondata[regione][0]["area"]]["seconda_dose"]
        reg_booster = data[jsondata[regione][0]["area"]]["booster"]

        for n_dosi in range(len(jsondata[regione])):        #iterazione su ogni oggetto nella lista di ogni regione
            reg_data_somministrazione.append(jsondata[regione][n_dosi]["data_somministrazione"][:10])
            reg_prima_dose.append(jsondata[regione][n_dosi]["prima_dose"])
            reg_seconda_dose.append(jsondata[regione][n_dosi]["seconda_dose"])
            reg_booster.append(jsondata[regione][n_dosi]["dose_addizionale_booster"])

    return data_sorted(data)


def getGithubUserContent():
    try:
        response = requests.get("https://raw.githubusercontent.com/italia/covid19-opendata-vaccini/master/dati/somministrazioni-vaccini-summary-latest.json")
        jsondata = json.loads(response.text)        #load della risposta response, caricata come testo e converita in json object
    except Exception as e:
        raise e

    rel = {}
    for data in jsondata["data"]:                   #iterazione sui dati di jsondata.data
        if not data["area"] in rel:
            rel[data["area"]] = []                  #creazione di una lista se la regione non è stata incontrata
        rel[data["area"]].append(data)              #aggiunta dei dati relativi alla regione

    with open("somministrazioni-vaccini-summary-latest.json","w") as wr:
        json.dump(rel,wr,indent=4)                                          #dump json sul file

def data_sorted(data):

    data_sorted = {}                #creazione dizionario sorted
    ita_data = {}                   #creazione dizionario italia
    ita_data["data"] = []
    ita_data["prima_dose"] = []
    ita_data["seconda_dose"] = []
    ita_data["booster"] = []
    ita_data["nome_area"] = "Italia"
    for regione in data:            #iterazione sulle regioni in data
        data_middleware = {}
        data_middleware1 = {}
        data_sorted[regione] = {}       #creazione di un dizionario data_sorted.regione

        for i in range(len(data[regione]["data"])):         #iterazione sulle date di somministrazione della regione in data
                                    #allocamento della data come chiave di un nuovo dizionario con value prima dose, seconda dose e booster
            data_middleware1[data[regione]["data"][i]] = {"prima_dose":data[regione]["prima_dose"][i],"seconda_dose":data[regione]["seconda_dose"][i],"booster":data[regione]["booster"][i]}

        data_middleware = {k: v for k, v in sorted(data_middleware1.items(), key=lambda item: item[0])}         #ordinamento del dizionario data_middleware1 basato sull'ordine crescente delle date

        data_sorted[regione]["nome_area"] = data[regione]["nome_area"]                                          #data_sorted immagazina i vari valori di data_middleware (ordinato)
        data_sorted[regione]["data"] = [i for i in data_middleware]
        data_sorted[regione]["prima_dose"] = [data_middleware[i]["prima_dose"] for i in data_middleware]
        data_sorted[regione]["seconda_dose"] = [data_middleware[i]["seconda_dose"] for i in data_middleware]
        data_sorted[regione]["booster"] = [data_middleware[i]["booster"] for i in data_middleware]



    for datetime in data_sorted["LAZ"]["data"]:                                 #iterazione sulle date di una regione con tutte le date di somministrazione
        n_dosi_ita_prima = 0
        n_dosi_ita_seconda = 0
        n_dosi_ita_booster = 0
        for regione in data_sorted:                                             #iterazione sulle regioni
            if datetime in data_sorted[regione]["data"]:                        #check se la data è presente nella regione
                i = data_sorted[regione]["data"].index(datetime)                #trovo l'index della data sulla lista delle date della regione
                n_dosi_ita_prima += data_sorted[regione]["prima_dose"][i]       #utilizzo l'index per aumentare le dosi dell'italia
                n_dosi_ita_seconda += data_sorted[regione]["seconda_dose"][i]
                n_dosi_ita_booster += data_sorted[regione]["booster"][i]
        ita_data["prima_dose"].append(n_dosi_ita_prima)                         #append del numero di dosi ogni data presente nella lista
        ita_data["seconda_dose"].append(n_dosi_ita_seconda)
        ita_data["booster"].append(n_dosi_ita_booster)

    ita_data["data"] = data_sorted["LAZ"]["data"]

    data_sorted["ITA"] = ita_data                                               #inserimento del dizionario in data_sorted


    return data_sorted


def export(data):
    try:
        workbook = xlsxwriter.Workbook('Vaccinazione.xlsx')     #creazione variabile workbook tramite libreria xlsxwriter
        for regione in data:                                    #iterazione sulle regioni in data
            worksheet = workbook.add_worksheet(regione)         #creazione di un nuovo worksheet con il nome della regione
            chart = workbook.add_chart({'type':'line'})         #aggiunta di un grafico chart line

            worksheet.write_column('A1',data[regione]["data"])              #il worksheet aggiunge i dati di data.regione
            worksheet.write_column('B1',data[regione]["prima_dose"])
            worksheet.write_column('C1',data[regione]["seconda_dose"])
            worksheet.write_column('D1',data[regione]["booster"])


                        #aggiunta dei dati con valori nelle colonne dalla B all D e dalla riga 1 alla n (numero di date presenti in data.regione)
            chart.add_series({'values':'=%s!$B$1:$B$%s'%(regione,len(data[regione]["data"])),'categories':'=%s!$A$1:$A$%s'%(regione,len(data[regione]["data"])),'line':{'color':'green'},'name':'Prima Dose'})
            chart.add_series({'values':'=%s!$C$1:$C$%s'%(regione,len(data[regione]["data"])),'categories':'=%s!$A$1:$A$%s'%(regione,len(data[regione]["data"])),'line':{'color':'red'},'name':'Seconda Dose'})
            chart.add_series({'values':'=%s!$D$1:$D$%s'%(regione,len(data[regione]["data"])),'categories':'=%s!$A$1:$A$%s'%(regione,len(data[regione]["data"])),'line':{'color':'blue'},'name':'Booster'})

            chart.set_size({'width':720,'height':570})

            chart.set_title({'name':data[regione]["nome_area"]})

            worksheet.insert_chart('F2',chart)      #inserimento del grafico chart nel worksheet
    except Exception as e:
        raise e
    finally:
        workbook.close()                            #in caso di successo chiusura di workbook e creazione di Vaccinazione.xlsx




if __name__ == '__main__':
    data = importer()
    export(data)
