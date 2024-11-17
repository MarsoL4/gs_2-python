import oracledb
import json
import pandas as pd
from datetime import datetime

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
    try:
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


def update_project():
    try:
        connection = conectarBD()
        if not connection:
            return
        cursor = connection.cursor()

        id_projeto = int(input("ID do projeto a ser atualizado: "))
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


def delete_project():
    try:
        connection = conectarBD()
        if not connection:
            return
        cursor = connection.cursor()

        id_projeto = int(input("ID do projeto a ser excluído: "))

        query = "DELETE FROM TBL_PROJETOS_SUSTENTAVEIS WHERE ID_PROJETO = :id_projeto"
        cursor.execute(query, {"id_projeto": id_projeto})
        connection.commit()
        print("\n🟢 Projeto excluído com sucesso!")
    except Exception as e:
        print(f"\n🔴 Erro ao excluir projeto: {e}")
    finally:
        close_connection(connection)


def query_projects():
    """
    Consulta os projetos no banco de dados e exibe os resultados de forma organizada.
    Retorna os resultados como uma lista de dicionários.
    """
    try:
        connection = conectarBD()
        if not connection:
            return []
        cursor = connection.cursor()

        status = input("Filtrar por status (ou deixe em branco para todos): ").strip()

        if status:
            query = """
                SELECT ID_PROJETO, DESCRICAO, CUSTO, STATUS, ID_TIPO_FONTE, ID_REGIAO
                FROM RM556310.TBL_PROJETOS_SUSTENTAVEIS
                WHERE UPPER(TRIM(STATUS)) = UPPER(TRIM(:status))
            """
            cursor.execute(query, {"status": status})
        else:
            query = """
                SELECT ID_PROJETO, DESCRICAO, CUSTO, STATUS, ID_TIPO_FONTE, ID_REGIAO
                FROM RM556310.TBL_PROJETOS_SUSTENTAVEIS
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
        
        return data  # Retorna os dados para exportação
    except Exception as e:
        print(f"\n🔴 Erro ao consultar projetos: {e}")
        return []
    finally:
        close_connection(connection)


# Exportação de Dados
def export_to_json(data, file_name=None):
    try:
        if not file_name:
            today = datetime.now().strftime("%Y-%m-%d")
            file_name = f"projetos_{today}.json"

        with open(file_name, "w", encoding="utf-8") as file:
            json.dump(data, file, indent=4, ensure_ascii=False)
        print(f"\n🟢 Dados exportados para o arquivo: {file_name}")
    except Exception as e:
        print(f"\n🔴 Erro ao exportar para JSON: {e}")


def export_to_excel(data, file_name=None):
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


# Menu Principal
def show_menu():
    print("\n=== MENU PRINCIPAL ===")
    print("1. Inserir Projeto")
    print("2. Atualizar Projeto")
    print("3. Excluir Projeto")
    print("4. Consultar Projetos")
    print("5. Exportar Dados")
    print("6. Sair")


def main():
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
            export_format = input("Escolha o formato (JSON/Excel): ").strip().lower()
            projects = query_projects()
            if projects:
                if export_format == "json":
                    export_to_json(projects)
                elif export_format == "excel":
                    export_to_excel(projects)
                else:
                    print("🔴 Formato inválido!")
            else:
                print("🔴 Nenhum dado disponível para exportação.")
        elif choice == "6":
            print("\n🟢 Saindo do sistema...")
            break
        else:
            print("\n🔴 Opção inválida. Tente novamente.")


if __name__ == "__main__":
    main()