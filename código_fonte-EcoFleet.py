import oracledb
import json
import pandas as pd
from datetime import datetime
import os


# Função para limpar o terminal
def clear_terminal():
    """
    Limpa o terminal para melhor organização visual.
    """
    os.system('cls' if os.name == 'nt' else 'clear')


# Conexão com o Banco de Dados Oracle
def conectarBD():
    """
    Estabelece conexão com o banco de dados Oracle.
    Retorna uma conexão ativa ou None em caso de erro.
    """
    try:
        conn = oracledb.connect(
            user="RM556310",
            password="130206",
            dsn="(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=oracle.fiap.com.br)(PORT=1521))(CONNECT_DATA=(SID=ORCL)))"
        )
        print("\n🟢 Conexão com o banco de dados estabelecida com sucesso!")
        return conn
    except oracledb.Error as e:
        print(f"\n🔴 Erro ao conectar ao banco de dados: {e}")
        return None


def close_connection(connection):
    """
    Encerra a conexão com o banco de dados.
    """
    if connection:
        connection.close()
        print("\n🟢 Conexão com o banco de dados encerrada.")


# Função para listar opções e selecionar
def listar_opcoes(tabela, campo_id, campo_nome):
    """
    Lista as opções de uma tabela e retorna o ID correspondente à escolha do usuário.
    """
    try:
        connection = conectarBD()
        if not connection:
            return None
        cursor = connection.cursor()

        query = f"SELECT {campo_id}, {campo_nome} FROM {tabela}"
        cursor.execute(query)
        results = cursor.fetchall()

        if results:
            print(f"\n=== Opções Disponíveis em {tabela} ===")
            for index, row in enumerate(results, start=1):
                print(f"{index}. {row[1]}")

            while True:
                try:
                    escolha = int(input(f"Escolha uma opção (1 a {len(results)}): "))
                    if 1 <= escolha <= len(results):
                        return results[escolha - 1][0]
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
        close_connection(connection)


# Validações
def validate_positive_number(value, field_name):
    """
    Valida se um valor é um número positivo.
    Solicita novamente a entrada até ser válido.
    """
    while True:
        try:
            number = float(value)
            if number <= 0:
                raise ValueError(f"O campo '{field_name}' deve ser um número positivo.")
            return number
        except ValueError as e:
            print(f"\n🔴 {e}")
            value = input(f"Insira novamente o campo '{field_name}': ")


def validate_non_empty_string(value, field_name):
    """
    Valida se um campo de entrada não está vazio.
    """
    while True:
        if not value.strip():
            print(f"🔴 O campo '{field_name}' não pode estar vazio.")
            value = input(f"Insira novamente o campo '{field_name}': ")
        else:
            return value.strip()


# Funções CRUD
def insert_project():
    """
    Insere um novo projeto no banco de dados.
    """
    try:
        clear_terminal()
        connection = conectarBD()
        if not connection:
            return
        cursor = connection.cursor()
        
        descricao = validate_non_empty_string(input("Descrição do projeto: "), "Descrição")
        custo = validate_positive_number(input("Custo do projeto: "), "Custo")
        status = validate_non_empty_string(input("Status do projeto: "), "Status")
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
        cursor.execute(query, {
            "descricao": descricao,
            "custo": custo,
            "status": status,
            "id_tipo_fonte": id_tipo_fonte,
            "id_regiao": id_regiao
        })
        connection.commit()
        print("\n🟢 Projeto inserido com sucesso!")
    except Exception as e:
        print(f"\n🔴 Erro ao inserir projeto: {e}")
    finally:
        close_connection(connection)
        input("\nPressione Enter para continuar...")


def update_project():
    """
    Atualiza um projeto existente no banco de dados.
    """
    try:
        clear_terminal()
        connection = conectarBD()
        if not connection:
            return
        cursor = connection.cursor()

        id_projeto = validate_positive_number(input("ID do projeto a ser atualizado: "), "ID do Projeto")
        descricao = validate_non_empty_string(input("Nova descrição do projeto: "), "Descrição")
        custo = validate_positive_number(input("Novo custo do projeto: "), "Custo")
        status = validate_non_empty_string(input("Novo status do projeto: "), "Status")

        query = """
            UPDATE TBL_PROJETOS_SUSTENTAVEIS
            SET DESCRICAO = :descricao, CUSTO = :custo, STATUS = :status
            WHERE ID_PROJETO = :id_projeto
        """
        cursor.execute(query, {
            "descricao": descricao,
            "custo": custo,
            "status": status,
            "id_projeto": id_projeto
        })
        connection.commit()
        print("\n🟢 Projeto atualizado com sucesso!")
    except Exception as e:
        print(f"\n🔴 Erro ao atualizar projeto: {e}")
    finally:
        close_connection(connection)
        input("\nPressione Enter para continuar...")


def delete_project():
    """
    Exclui um projeto do banco de dados.
    """
    try:
        clear_terminal()
        connection = conectarBD()
        if not connection:
            return
        cursor = connection.cursor()

        id_projeto = validate_positive_number(input("ID do projeto a ser excluído: "), "ID do Projeto")

        query = "DELETE FROM TBL_PROJETOS_SUSTENTAVEIS WHERE ID_PROJETO = :id_projeto"
        cursor.execute(query, {"id_projeto": id_projeto})
        connection.commit()
        print("\n🟢 Projeto excluído com sucesso!")
    except Exception as e:
        print(f"\n🔴 Erro ao excluir projeto: {e}")
    finally:
        close_connection(connection)
        input("\nPressione Enter para continuar...")


def query_projects():
    """
    Consulta os projetos no banco de dados e exibe os resultados de forma organizada.
    Retorna os resultados como uma lista de dicionários.
    """
    try:
        clear_terminal()
        connection = conectarBD()
        if not connection:
            return []
        cursor = connection.cursor()

        print("\n=== Filtrar Projetos ===")
        print("1. Todos os projetos")
        print("2. Apenas os projetos em andamento")
        print("3. Apenas os projetos concluídos")
        while True:
            try:
                filter_choice = int(input("Escolha uma opção (1-3): "))
                if filter_choice not in [1, 2, 3]:
                    print("🔴 Opção inválida. Tente novamente.")
                    continue
                break
            except ValueError:
                print("🔴 Entrada inválida. Por favor, insira um número.")

        if filter_choice == 1:
            query = """
                SELECT ID_PROJETO, DESCRICAO, CUSTO, STATUS, ID_TIPO_FONTE, ID_REGIAO
                FROM RM556310.TBL_PROJETOS_SUSTENTAVEIS
            """
            cursor.execute(query)
        elif filter_choice == 2:
            query = """
                SELECT ID_PROJETO, DESCRICAO, CUSTO, STATUS, ID_TIPO_FONTE, ID_REGIAO
                FROM RM556310.TBL_PROJETOS_SUSTENTAVEIS
                WHERE STATUS = 'Em andamento'
            """
            cursor.execute(query)
        else:
            query = """
                SELECT ID_PROJETO, DESCRICAO, CUSTO, STATUS, ID_TIPO_FONTE, ID_REGIAO
                FROM RM556310.TBL_PROJETOS_SUSTENTAVEIS
                WHERE STATUS = 'Concluído'
            """
            cursor.execute(query)

        results = cursor.fetchall()

        # Converte os resultados para uma lista de dicionários
        data = [
            {
                "ID_PROJETO": row[0],
                "DESCRICAO": row[1],
                "CUSTO": row[2],
                "STATUS": row[3],
                "ID_TIPO_FONTE": row[4],
                "ID_REGIAO": row[5]
            }
            for row in results
        ]

        if results:
            print("\n🟢 Projetos encontrados:")
            for project in data:
                print("\n-------------------------------")
                print(f"ID do Projeto: {project['ID_PROJETO']}")
                print(f"Descrição: {project['DESCRICAO']}")
                print(f"Custo: R${project['CUSTO']:,.2f}")
                print(f"Status: {project['STATUS']}")
                print(f"ID do Tipo de Fonte: {project['ID_TIPO_FONTE']}")
                print(f"ID da Região: {project['ID_REGIAO']}")
                print("-------------------------------")
        else:
            print("\n🔴 Nenhum projeto encontrado.")

        input("\nPressione Enter para continuar...")
        return data  # Retorna os dados para exportação
    except Exception as e:
        print(f"\n🔴 Erro ao consultar projetos: {e}")
        return []
    finally:
        close_connection(connection)


# Exportação de Dados
def export_to_json(data, file_name=None):
    """
    Exporta os dados para um arquivo JSON.
    """
    try:
        if not file_name:
            today = datetime.now().strftime("%Y-%m-%d")
            file_name = f"projetos_{today}.json"

        with open(file_name, "w", encoding="utf-8") as file:
            json.dump(data, file, indent=4, ensure_ascii=False)
        print(f"\n🟢 Dados exportados para o arquivo: {file_name}")
    except Exception as e:
        print(f"\n🔴 Erro ao exportar para JSON: {e}")
    input("\nPressione Enter para continuar...")


def export_to_excel(data, file_name=None):
    """
    Exporta os dados para um arquivo Excel.
    """
    try:
        if not file_name:
            today = datetime.now().strftime("%Y-%m-%d")
            file_name = f"projetos_{today}.xlsx"

        # Especificando o engine openpyxl
        df = pd.DataFrame(data, columns=["ID_PROJETO", "DESCRICAO", "CUSTO", "STATUS", "ID_TIPO_FONTE", "ID_REGIAO"])
        df.to_excel(file_name, index=False, engine="openpyxl")
        print(f"\n🟢 Dados exportados para o arquivo: {file_name}")
    except ModuleNotFoundError:
        print("\n🔴 O módulo 'openpyxl' não está instalado. Por favor, instale-o usando 'pip install openpyxl'.")
    except Exception as e:
        print(f"\n🔴 Erro ao exportar para Excel: {e}")
    input("\nPressione Enter para continuar...")


# Menu Principal
def show_menu():
    """
    Exibe o menu principal.
    """
    clear_terminal()
    print("\n=== MENU PRINCIPAL ===")
    print("1. Inserir Projeto")
    print("2. Atualizar Projeto")
    print("3. Excluir Projeto")
    print("4. Consultar Projetos")
    print("5. Exportar Dados")
    print("6. Sair")


def main():
    """
    Função principal que controla o fluxo do programa.
    """
    while True:
        show_menu()
        choice = input("Escolha uma opção: ")
        if choice == "1":
            insert_project()
        elif choice == "2":
            update_project()
        elif choice == "3":
            delete_project()
        elif choice == "4":
            query_projects()
        elif choice == "5":
            projects = query_projects()
            if projects:
                while True:
                    print("\n=== Exportar Dados ===")
                    print("1. Exportar para JSON")
                    print("2. Exportar para Excel")
                    export_choice = input("Escolha uma opção (1-2): ")
                    if export_choice == "1":
                        export_to_json(projects)
                        break
                    elif export_choice == "2":
                        export_to_excel(projects)
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


if __name__ == "__main__":
    main()