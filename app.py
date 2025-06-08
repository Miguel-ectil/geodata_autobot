import requests
import sqlite3
from openpyxl import Workbook
from datetime import datetime

# === Função para buscar dados do país na API ===
def buscar_dados_pais(nome_pais):
    url = f"https://restcountries.com/v3.1/name/{nome_pais}"
    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()[0]

        pais = {
            "nome_comum": data.get("name", {}).get("common", "N/A"),
            "nome_oficial": data.get("name", {}).get("official", "N/A"),
            "capital": data.get("capital", ["N/A"])[0],
            "continente": data.get("continents", ["N/A"])[0],
            "regiao": data.get("region", "N/A"),
            "sub_regiao": data.get("subregion", "N/A"),
            "populacao": data.get("population", "N/A"),
            "area": data.get("area", "N/A"),
            "moeda_nome": list(data.get("currencies", {}).values())[0].get("name", "N/A"),
            "moeda_simbolo": list(data.get("currencies", {}).values())[0].get("symbol", "N/A"),
            "idioma_principal": list(data.get("languages", {}).values())[0],
            "fuso_horario": data.get("timezones", ["N/A"])[0],
            "bandeira_url": data.get("flags", {}).get("png", "N/A"),
        }

        return pais

    except Exception as e:
        print(f"Erro ao buscar o país '{nome_pais}': {e}")
        return None

# === Função para armazenar no banco de dados ===
def armazenar_dados(paises):
    conn = sqlite3.connect("paises.db")
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS paises (
            nome_comum TEXT,
            nome_oficial TEXT,
            capital TEXT,
            continente TEXT,
            regiao TEXT,
            sub_regiao TEXT,
            populacao INTEGER,
            area REAL,
            moeda_nome TEXT,
            moeda_simbolo TEXT,
            idioma_principal TEXT,
            fuso_horario TEXT,
            bandeira_url TEXT
        )
    """)

    for pais in paises:
        cursor.execute("""
            INSERT INTO paises VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, tuple(pais.values()))

    conn.commit()
    conn.close()

# === Função para gerar relatório Excel ===
def gerar_relatorio_excel(paises, nome_aluno):
    wb = Workbook()
    ws = wb.active
    ws.title = "Relatório de Países"

    cabecalhos = [
        "Nome Comum", "Nome Oficial", "Capital", "Continente", "Região", "Sub-Região",
        "População", "Área (km²)", "Moeda", "Símbolo", "Idioma", "Fuso Horário", "URL da Bandeira"
    ]
    ws.append(cabecalhos)

    for pais in paises:
        linha = [
            pais["nome_comum"],
            pais["nome_oficial"],
            pais["capital"],
            pais["continente"],
            pais["regiao"],
            pais["sub_regiao"],
            pais["populacao"],
            pais["area"],
            pais["moeda_nome"],
            pais["moeda_simbolo"],
            pais["idioma_principal"],
            pais["fuso_horario"],
            pais["bandeira_url"],
        ]
        ws.append(linha)

    arquivo = f"relatorio_{nome_aluno.lower().replace(' ', '_')}.xlsx"
    wb.save(arquivo)
    print(f"Relatório salvo como: {arquivo}")

# === Função principal ===
def main():
    nome_aluno = input("Digite seu nome completo: ")
    paises_input = input("Digite três países separados por vírgula: ").split(",")

    paises_input = [pais.strip() for pais in paises_input[:3]] 
    dados_paises = []

    for nome in paises_input:
        dados = buscar_dados_pais(nome)
        if dados:
            dados_paises.append(dados)

    if dados_paises:
        armazenar_dados(dados_paises)
        gerar_relatorio_excel(dados_paises, nome_aluno)
    else:
        print("Nenhum dado válido coletado.")

if __name__ == "__main__":
    main()
