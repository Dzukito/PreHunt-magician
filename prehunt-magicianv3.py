import requests
from bs4 import BeautifulSoup
import csv
import openpyxl
from openpyxl import load_workbook


lista_grupos = []
dic_grupo_tec = {}
i = 0
dic_id_grupo = {}
dic_tecnicas_nombre = {}
dic_tecnicas_tacticas = {}
dic_tecnicas_description = {}
dic_tecnicas_plataforma = {}
dic_tecnicas_data={}


# Obtener actores seleccionados del Excel
# Ruta del archivo de Excel
archivo_excel = "PrehuntVacio.xlsx"

# Cargar el archivo de Excel
libro = openpyxl.load_workbook(archivo_excel)

# Seleccionar la hoja de cálculo (sheet) específica
hoja = libro["Actores seleccionados"]

# Leer el valor de una celda específica
actores= []
actores.append(str(hoja["B8"].value))
actores.append(hoja["B9"].value)
actores.append(hoja["B10"].value)
actores.append(hoja["B11"].value)

# ---------------------- RECORRER MITRE ----------------------------------------


url = 'https://attack.mitre.org/groups/'

# Realizar la solicitud HTTP GET a la página
response = requests.get(url)

# Verificar si la solicitud fue exitosa
if response.status_code == 200:
    # Crear un objeto BeautifulSoup con el contenido HTML de la página
    soup = BeautifulSoup(response.content, 'html.parser')

    # Encontrar la tabla que contiene los grupos
    table = soup.find('table')

    # Recorrer las filas de la tabla
    for row in table.find_all('tr'):
        # Encontrar la celda que contiene el ID del grupo
        id_cell = row.find('td')

        # Extraer el ID del grupo si se encuentra la celda
        if id_cell:
            group_id = id_cell.text.strip()
            if group_id.startswith('G'):  # Filtrar solo los IDs que comienzan con 'G'
                lista_grupos.append(group_id)

    for row_grupos in table.find_all('tr'):
        # Encontrar todas las celdas de la fila
        cells_grupos = row_grupos.find_all('td')
        if len(cells_grupos) >= 2:  # Verificar si hay al menos dos celdas
            nombre_grupo = cells_grupos[1].text.strip()
            dic_id_grupo[lista_grupos[i]] = [nombre_grupo]  # {G123: [nombre_grupo]}
            i = i + 1

    


    lista_id_seleccionados=[]
    for actor_id, actor_nombre in dic_id_grupo.items():  
      if actor_nombre[0] in actores:
        lista_id_seleccionados.append(actor_id)

    print(lista_id_seleccionados)
    if len(lista_id_seleccionados)<4:
        print("Revisar actores, falto alguno")
        


      
    for grupo in lista_id_seleccionados:
        url_grupo = url + "/" + grupo
        response_grupo = requests.get(url_grupo)
        soup_grupo = BeautifulSoup(response_grupo.content, 'html.parser')

        table_tecnicas = soup_grupo.find('table', {'class': 'table techniques-used background table-bordered'})

        # Recorrer las filas de la tabla
        lista_id = []
        try:
            tec_id_aux=""
            for row_tec in table_tecnicas.find_all('tr'):
                # Encontrar todas las celdas de la fila
                cells = row_tec.find_all('td')
                if len(cells) >= 2:  # Verificar si hay al menos dos celdas
                    tec_id = cells[1].text.strip()
                    if tec_id.startswith('T'):  # Filtrar solo los IDs que comienzan con 'T'
                        tec_id_aux=tec_id
                        subtec_id = cells[2].text.strip()  # Obtener el ID de la subtecnología si existe
                        if subtec_id.startswith('.'):
                            tec_id += subtec_id  # Concatenar el ID de la subtecnología a la tecnología principal
                        lista_id.append(tec_id) 
                    else:
                       subtec_id = cells[2].text.strip()  # Obtener el ID de la subtecnología si existe
                       if subtec_id.startswith('.'):
                        lista_id.append(tec_id_aux + subtec_id) # Concatenar el ID de la subtecnología a la tecnología principal
                        

            dic_grupo_tec[grupo] = lista_id  
            lista_id = []
        except:
            dic_grupo_tec[grupo] = []  # Grupo sin técnicas
        # {GXXX: [tec_id, tec_id, ...]} Resultado final
    
    
    for grupo, tecnicas in dic_grupo_tec.items():
      for tecnica in tecnicas:
        if tecnica not in dic_tecnicas_nombre:
          url_id = 'https://attack.mitre.org/techniques/'

          if '.' in tecnica:
            subtecnica_id = tecnica.split('.')
            url_id += subtecnica_id[0]+'/'+subtecnica_id[1]
          else:
            url_id += tecnica

          # Realizar la solicitud HTTP GET a la página
          response_id = requests.get(url_id)
          print(url_id)
          # Verificar si la solicitud fue exitosa
          if response_id.status_code == 200:
              # Crear un objeto BeautifulSoup con el contenido HTML de la página
              soup_id = BeautifulSoup(response_id.content, 'html.parser')

              # Titulo
              titulo = soup_id.find('h1').text.strip() #, {'id': 'table techniques-used background table-bordered'}
              if ':' in titulo:
                titulo_separado= titulo.split(":")
                titulo= titulo_separado[0]+": "+titulo_separado[1].lstrip()
              
              dic_tecnicas_nombre[tecnica]=tecnica+' - '+titulo #{T1234: T1234-Blablabla, ...}
              
              #Tacticas
              tactic_element = soup_id.find(id='card-tactics')
              tactic_links = tactic_element.find_all('a')
              tactic_data = []
              for link in tactic_links:
                  tactic_id = link['href'].split('/')[-1]
                  tactic_value = link.text.strip()
                  tactic_data.append(tactic_id + ' - ' + tactic_value) #(f'{tactic_id} - {tactic_value}')
             
              if(tecnica not in dic_tecnicas_tacticas):
                dic_tecnicas_tacticas[tecnica]=tactic_data #{T1234: [...], ...}

              #Descripcion
              description_element = soup_id.find(class_='description-body')
              description_text = description_element.get_text(strip=True)
              #print(description_text)
              if(tecnica not in dic_tecnicas_description):
                dic_tecnicas_description[tecnica]=description_text

              #Plataformas
              divs_plataformas = soup_id.find_all('div', class_='row card-data')
              # Iterar sobre los divs de las plataformas
              for div_plataformas in divs_plataformas:
                  # Encontrar el span que contiene el título 'Platforms:'
                  span_titulo = div_plataformas.find('span', class_='h5 card-title')
               
                  if span_titulo and span_titulo.text.strip() == 'Platforms:':
                      # Obtener el valor de las plataformas
                      plataformas = div_plataformas.find('div', class_='col-md-11 pl-0').text.strip() #Platforms: Azure AD, Google Workspace, IaaS, Linux, Office 365, SaaS, Windows, macOS
                      lista_plataformas=plataformas[11:].split(',')

                      if(tecnica not in dic_tecnicas_plataforma):
                        dic_tecnicas_plataforma[tecnica]=lista_plataformas
                    

              # Data Sources
              lista_data = []
              try:
                data_source_aux=""
                table = soup_id.find('table', class_='table datasources-table table-bordered')
                # Encuentra todas las filas de datos en el cuerpo de la tabla
                rows = table.find_all('tr')

                # Itera sobre cada fila (omitimos la primera fila de encabezados)
                for row in rows[1:]:
                    # Encuentra los elementos td de la fila
                    tds = row.find_all('td')
                    
                    # Extrae los valores de las columnas "Data Source" y "Data Component"
                    data_source = tds[1].text.strip()
                    if data_source !='':
                      data_source_aux=data_source
                      data_component = tds[2].text.strip()
                      lista_data.append(f"{data_source} : {data_component}")
                    else:
                      data_component = tds[2].text.strip()
                      lista_data.append(f"{data_source_aux} : {data_component}")

                if(tecnica not in dic_tecnicas_data):
                  dic_tecnicas_data[tecnica]=lista_data
                
              except:    
                title_element = soup_id.find('h2', class_='pt-3', id='detection')
                text = title_element.find_next('div').text.strip()
                if(tecnica not in dic_tecnicas_data):
                  dic_tecnicas_data[tecnica]=text
              
              

                              

              

    # output
    hoja_ttps_de_actores = libro["TTPs de Actores"]
    fila=8;
  
    for grupo, tecnicas in dic_grupo_tec.items():
          for tecnica in tecnicas:
            hoja_ttps_de_actores["B"+str(fila)] = tecnica
            hoja_ttps_de_actores["C"+str(fila)] = dic_id_grupo[grupo][0]
            hoja_ttps_de_actores["D"+str(fila)] = str(dic_tecnicas_tacticas[tecnica]).replace("[","").replace("]","").replace("'","")
            hoja_ttps_de_actores["E"+str(fila)] = dic_tecnicas_nombre[tecnica]
            hoja_ttps_de_actores["F"+str(fila)] = dic_tecnicas_description[tecnica]
            fila+=1

    hoja_visibilidad = libro["Visibilidad por Técnicas"]
    fila=8;
  
    for grupo, tecnicas in dic_grupo_tec.items():
          for tecnica in tecnicas:
            hoja_visibilidad["C"+str(fila)] = tecnica
            hoja_visibilidad["D"+str(fila)] = dic_tecnicas_nombre[tecnica]
            hoja_visibilidad["E"+str(fila)] = str(dic_tecnicas_tacticas[tecnica]).replace("[","").replace("]","").replace("'","")
            hoja_visibilidad["G"+str(fila)] = dic_tecnicas_description[tecnica]
            hoja_visibilidad["H"+str(fila)] = str(dic_tecnicas_plataforma[tecnica]).replace("[","").replace("]","").replace("'","")
            hoja_visibilidad["I"+str(fila)] = str(dic_tecnicas_data[tecnica]).replace("[","").replace("]","").replace("'","")
            fila+=1
    libro.save(archivo_excel)
            

    with open('output.csv', 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['ID de la Tecnica o Subtecnica', 'Actor', 'Tactica', 'Tecnica o Subtecnica','Descripcion de la Tecnica o Subtecnica (MITRE)', 'Plataformas', 'Data Sources'])

        # Escribir los datos en el archivo CSV
        for grupo, tecnicas in dic_grupo_tec.items():
            for tecnica in tecnicas:
                writer.writerow([tecnica, dic_id_grupo[grupo][0],str(dic_tecnicas_tacticas[tecnica]).replace("[","").replace("]","").replace("'",""),dic_tecnicas_nombre[tecnica],dic_tecnicas_description[tecnica], dic_tecnicas_plataforma[tecnica],dic_tecnicas_data[tecnica]])

else:
    print('No se pudo acceder a la página:', response.status_code)
