import oracledb
import json
import pandas as pd
from datetime import datetime
import os


# Função para limpar o terminal
def limpar_terminal():
    """
    Limpa o terminal para melhor organização visual.
    """
    os.system("cls" if os.name == "nt" else "clear")


# Conexão com o Banco de Dados Oracle
def conectarBD():
    """
    Estabelece conexão com o banco de dados Oracle.

    Retorna:
        Uma conexão ativa ou None em caso de erro.
    """
    try:
        conn = oracledb.connect(
            user="RM556310",
            password="130206",
            dsn="(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=oracle.fiap.com.br)(PORT=1521))(CONNECT_DATA=(SID=ORCL)))",
        )
        return conn
    except oracledb.Error as e:
        print(f"\n🔴 Erro ao conectar ao banco de dados: {e}")
        return None


# Fechar a conexão com o banco
def fechar_conexao(conexao):
    """
    Encerra a conexão com o banco de dados.

    Parâmetros:
        conexao: Objeto de conexão com o banco de dados.
    """
    if conexao:
        conexao.close()


# Listar opções de tabelas
def listar_opcoes(tabela, campo_id, campo_nome):
    """
    Lista as opções de uma tabela e retorna o ID correspondente à escolha do usuário.

    Parâmetros:
        tabela (str): Nome da tabela.
        campo_id (str): Nome do campo ID da tabela.
        campo_nome (str): Nome do campo de descrição da tabela.

    Retorna:
        int: ID selecionado pelo usuário ou None em caso de erro.
    """
    try:
        conexao = conectarBD()
        if not conexao:
            return None
        cursor = conexao.cursor()

        query = f"SELECT {campo_id}, {campo_nome} FROM {tabela}"
        cursor.execute(query)
        resultados = cursor.fetchall()

        if resultados:
            print(f"\n=== Opções disponíveis em {tabela} ===")
            for index, linha in enumerate(resultados, start=1):
                print(f"{index}. {linha[1]}")

            while True:
                try:
                    escolha = int(input(f"Escolha uma opção (1 a {len(resultados)}): "))
                    if 1 <= escolha <= len(resultados):
                        return resultados[escolha - 1][0]
                    else:
                        print("🔴 Opção inválida. Tente novamente.")
                except ValueError:
                    print("🔴 Entrada inválida. Por favor, insira um número.")
        else:
            print(f"🔴 Nenhum dado encontrado na tabela {tabela}.")
            return None
    except Exception as e:
        print(f"\n🔴 Erro ao listar opções na tabela {tabela}: {e}")
        return None
    finally:
        fechar_conexao(conexao)


# Validação de números positivos
def validar_numero_positivo(valor, nome_campo):
    """
    Valida se um valor é um número positivo.

    Parâmetros:
        valor (str): Entrada do usuário.
        nome_campo (str): Nome do campo a ser validado.

    Retorna:
        float: Número positivo validado.
    """
    while True:
        try:
            numero = float(valor)
            if numero <= 0:
                raise ValueError(f"O campo '{nome_campo}' deve ser um número positivo.")
            return numero
        except ValueError as e:
            print(f"\n🔴 {e}")
            valor = input(f"Insira novamente o campo '{nome_campo}': ")


# Validação de strings não vazias
def validar_string_nao_vazia(valor, nome_campo):
    """
    Valida se uma string não está vazia.

    Parâmetros:
        valor (str): Entrada do usuário.
        nome_campo (str): Nome do campo a ser validado.

    Retorna:
        str: String validada.
    """
    while True:
        if not valor.strip():
            print(f"🔴 O campo '{nome_campo}' não pode estar vazio.")
            valor = input(f"Insira novamente o campo '{nome_campo}': ")
        else:
            return valor.strip()


# Inserir um novo projeto
def inserir_projeto():
    """
    Insere um novo projeto no banco de dados.
    """
    try:
        limpar_terminal()
        print("\n=== Cadastrando um novo projeto ===")
        conexao = conectarBD()
        if not conexao:
            return
        cursor = conexao.cursor()

        descricao = validar_string_nao_vazia(input("Descrição do projeto: "), "Descrição")
        custo = validar_numero_positivo(input("Custo do projeto: "), "Custo")
        status = validar_string_nao_vazia(input("Status do projeto: "), "Status")
        id_tipo_fonte = listar_opcoes("TBL_TIPO_FONTES", "ID_TIPO_FONTE", "NOME")
        id_regiao = listar_opcoes("TBL_REGIOES_SUSTENTAVEIS", "ID_REGIAO", "NOME")

        if id_tipo_fonte is None or id_regiao is None:
            print("🔴 Operação cancelada devido a falha na seleção de dados.")
            return

        query = """
            INSERT INTO TBL_PROJETOS_SUSTENTAVEIS 
            (DESCRICAO, CUSTO, STATUS, ID_TIPO_FONTE, ID_REGIAO)
            VALUES (:descricao, :custo, :status, :id_tipo_fonte, :id_regiao)
        """
        cursor.execute(
            query,
            {
                "descricao": descricao,
                "custo": custo,
                "status": status,
                "id_tipo_fonte": id_tipo_fonte,
                "id_regiao": id_regiao,
            },
        )
        conexao.commit()
        print("\n🟢 Projeto inserido com sucesso!")
    except Exception as e:
        print(f"\n🔴 Erro ao inserir projeto: {e}")
    finally:
        fechar_conexao(conexao)
        input("\nPressione Enter para continuar...")


# Atualizar um projeto existente
def atualizar_projeto():
    """
    Atualiza um projeto existente no banco de dados.
    """
    try:
        limpar_terminal()
        print("\n=== Atualizando um projeto ===")
        conexao = conectarBD()
        if not conexao:
            return
        cursor = conexao.cursor()

        id_projeto = validar_numero_positivo(input("ID do projeto a ser atualizado: "), "ID do Projeto")
        descricao = validar_string_nao_vazia(input("Nova descrição do projeto: "), "Descrição")
        custo = validar_numero_positivo(input("Novo custo do projeto: "), "Custo")
        status = validar_string_nao_vazia(input("Novo status do projeto: "), "Status")

        query = """
            UPDATE TBL_PROJETOS_SUSTENTAVEIS
            SET DESCRICAO = :descricao, CUSTO = :custo, STATUS = :status
            WHERE ID_PROJETO = :id_projeto
        """
        cursor.execute(
            query,
            {"descricao": descricao, "custo": custo, "status": status, "id_projeto": id_projeto},
        )
        conexao.commit()
        print("\n🟢 Projeto atualizado com sucesso!")
    except Exception as e:
        print(f"\n🔴 Erro ao atualizar projeto: {e}")
    finally:
        fechar_conexao(conexao)
        input("\nPressione Enter para continuar...")

# Excluir um projeto
def excluir_projeto():
    """
    Exclui um projeto do banco de dados.
    """
    try:
        limpar_terminal()
        print("\n=== Excluindo um projeto ===")
        conexao = conectarBD()
        if not conexao:
            return
        cursor = conexao.cursor()

        id_projeto = validar_numero_positivo(input("ID do projeto a ser excluído: "), "ID do Projeto")

        query = "DELETE FROM TBL_PROJETOS_SUSTENTAVEIS WHERE ID_PROJETO = :id_projeto"
        cursor.execute(query, {"id_projeto": id_projeto})
        conexao.commit()
        print("\n🟢 Projeto excluído com sucesso!")
    except Exception as e:
        print(f"\n🔴 Erro ao excluir projeto: {e}")
    finally:
        fechar_conexao(conexao)
        input("\nPressione Enter para continuar...")


# Consultar projetos
def consultar_projetos(para_exportar=False) -> list:
    """
    Consulta os projetos no banco de dados e exibe os resultados de forma organizada.

    Parâmetros:
        para_exportar (bool): Define se o texto deve ser ajustado para exportação ou consulta.

    Retorna:
        list: Lista de dicionários com os dados dos projetos.
    """
    try:
        if para_exportar:
            print("\n=== Selecione os projetos que deseja exportar ===")
        else:
            limpar_terminal()
            print("\n=== Consultando projetos ===")
        
        print("1. Todos os projetos")
        print("2. Apenas os projetos em andamento")
        print("3. Apenas os projetos concluídos")

        escolha = input("Escolha uma opção (1-3): ").strip()
        if escolha not in ["1", "2", "3"]:
            print("\n🔴 Opção inválida.")
            return []

        consulta = (
            "SELECT ID_PROJETO, DESCRICAO, CUSTO, STATUS, ID_TIPO_FONTE, ID_REGIAO "
            "FROM RM556310.TBL_PROJETOS_SUSTENTAVEIS"
        )
        if escolha == "2":
            consulta += " WHERE STATUS = 'Em andamento'"
        elif escolha == "3":
            consulta += " WHERE STATUS = 'Concluído'"

        conexao = conectarBD()
        if not conexao:
            return []
        cursor = conexao.cursor()
        cursor.execute(consulta)

        resultados = cursor.fetchall()
        if not resultados:
            print("\n🔴 Nenhum projeto encontrado.")
        else:
            if para_exportar:
                print("\n=== Projetos que serão exportados ===")
            for projeto in resultados:
                print(f"\nID: {projeto[0]} | Descrição: {projeto[1]} | Custo: R${projeto[2]:,.2f} | Status: {projeto[3]}")

        if not para_exportar:
            input("\nPressione Enter para continuar...")

        return resultados
    finally:
        fechar_conexao(conexao)

# Exportação de Dados para JSON
def exportar_json(dados, nome_arquivo=None):
    """
    Exporta os dados para um arquivo JSON de forma organizada e legível.

    Parâmetros:
        dados (list): Lista de dicionários com os dados a serem exportados.
        nome_arquivo (str): Nome do arquivo para exportação. Se None, será gerado automaticamente.
    """
    try:
        if not nome_arquivo:
            hoje = datetime.now().strftime("%Y-%m-%d")
            nome_arquivo = f"projetos_{hoje}.json"

        # Garantindo que os dados sejam uma lista de dicionários
        if isinstance(dados, list) and all(isinstance(item, (list, tuple)) for item in dados):
            dados = [
                {
                    "ID_PROJETO": item[0],
                    "DESCRICAO": item[1],
                    "CUSTO": item[2],
                    "STATUS": item[3],
                    "ID_TIPO_FONTE": item[4],
                    "ID_REGIAO": item[5]
                }
                for item in dados
            ]

        with open(nome_arquivo, "w", encoding="utf-8") as arquivo:
            json.dump(dados, arquivo, indent=4, ensure_ascii=False)
        print(f"\n🟢 Dados exportados para o arquivo: {nome_arquivo}")
    except Exception as e:
        print(f"\n🔴 Erro ao exportar para JSON: {e}")
    input("\nPressione Enter para continuar...")


# Exportação de Dados para Excel
def exportar_DataFrame(dados, nome_arquivo=None):
    """
    Exporta os dados para um DataFrame (arquivo .xlsx).

    Parâmetros:
        dados (list): Lista de dicionários com os dados a serem exportados.
        nome_arquivo (str): Nome do arquivo para exportação. Se None, será gerado automaticamente.
    """
    try:
        if not nome_arquivo:
            hoje = datetime.now().strftime("%Y-%m-%d")
            nome_arquivo = f"projetos_{hoje}.xlsx"

        df = pd.DataFrame(
            dados,
            columns=["ID_PROJETO", "DESCRICAO", "CUSTO", "STATUS", "ID_TIPO_FONTE", "ID_REGIAO"],
        )
        df.to_excel(nome_arquivo, index=False, engine="openpyxl")
        print(f"\n🟢 Dados exportados para o arquivo: {nome_arquivo}")
    except ModuleNotFoundError:
        print("\n🔴 O módulo 'openpyxl' não está instalado. Por favor, instale-o usando 'pip install openpyxl'.")
    except Exception as e:
        print(f"\n🔴 Erro ao exportar para Excel: {e}")
    input("\nPressione Enter para continuar...")


# Menu Principal
def exibir_menu():
    """
    Exibe o menu principal.
    """
    limpar_terminal()
    print("\n=== MENU PRINCIPAL ===")
    print("1. Cadastrar novo Projeto")
    print("2. Atualizar Projeto Existente")
    print("3. Excluir Projeto pelo ID")
    print("4. Consultar Projetos pelo status")
    print("5. Exportar Projetos para JSON ou DataFrame")
    print("6. Sair")


# Função Principal
def main():
    """
    Função principal que controla o fluxo do programa.
    """
    while True:
        exibir_menu()
        choice = input("Escolha uma opção: ")
        if choice == "1":
            inserir_projeto()
        elif choice == "2":
            atualizar_projeto()
        elif choice == "3":
            excluir_projeto()
        elif choice == "4":
            consultar_projetos()  # Consulta normal
        elif choice == "5":
            projects = consultar_projetos(para_exportar=True)  # Ajuste para exportação
            if projects:
                while True:
                    print("\n=== Exportar Dados ===")
                    print("1. Exportar para JSON")
                    print("2. Exportar para Excel")
                    export_choice = input("Escolha uma opção (1-2): ")
                    if export_choice == "1":
                        exportar_json(projects)
                        break
                    elif export_choice == "2":
                        exportar_DataFrame(projects)
                        break
                    else:
                        print("🔴 Opção inválida. Tente novamente.")
            else:
                print("🔴 Nenhum dado disponível para exportação.")
        elif choice == "6":
            print("\n🟢 Saindo do sistema...")
            break
        else:
            print("\n🔴 Opção inválida. Tente novamente.")
            input("\nPressione Enter para continuar...")

# Executa o programa
main()