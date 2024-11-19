# ================================ Imports ================================ 
import oracledb
import json
import pandas as pd
from datetime import datetime
import os


# ============================== Subalgoritmos ============================

# Limpa o terminal para uma exibição mais limpa
def limpar_terminal() -> None:
    # Determina o comando para limpar o terminal dependendo do sistema operacional
    os.system("cls" if os.name == "nt" else "clear")

# Estabelece conexão com o Banco de Dados
def conectarBD() -> oracledb.Connection | None:
    try:
        # Conecta ao banco de dados usando as credenciais fornecidas
        conn = oracledb.connect(
            user="RM556310",
            password="130206",
            dsn="(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=oracle.fiap.com.br)(PORT=1521))(CONNECT_DATA=(SID=ORCL)))",
        )
        return conn  # Retorna a conexão ativa
    except oracledb.Error as e:
        # Exibe mensagem de erro caso a conexão falhe
        print(f"\n🔴 Erro ao conectar ao banco de dados: {e}")
        return None

# Encerra a conexão com o banco de dados
def fechar_conexao(conexao: oracledb.Connection) -> None:
    if conexao:
        conexao.close()  # Fecha a conexão com o Banco de Dados

# Lista as opções de uma tabela do banco de dados com mensagens personalizadas
def listar_opcoes(tabela: str, campo_id: str, campo_nome: str) -> int | None:
    try:
        conexao = conectarBD()  # Abre uma conexão com o banco
        if not conexao:
            return None
        cursor = conexao.cursor()

        # Monta e executa a query para buscar os dados da tabela
        query = f"SELECT {campo_id}, {campo_nome} FROM {tabela}"
        cursor.execute(query)
        resultados = cursor.fetchall()  # Armazena os resultados

        if resultados:
            # Define títulos amigáveis ao usuário
            if tabela == "TBL_TIPO_FONTES":
                print("\n=== Selecione o tipo de fonte ===")
            elif tabela == "TBL_REGIOES_SUSTENTAVEIS":
                print("\n=== Selecione a região sustentável ===")
            else:
                print(f"\n=== Opções disponíveis em {tabela} ===")

            # Exibe os resultados formatados como opções para o usuário
            for index, linha in enumerate(resultados, start=1):
                print(f"{index}. {linha[1]}")

            # Solicita que o usuário escolha uma opção válida
            while True:
                try:
                    escolha = int(input(f"Escolha uma opção (1 a {len(resultados)}): "))
                    if 1 <= escolha <= len(resultados):
                        return resultados[escolha - 1][0]  # Retorna o ID da opção escolhida
                    else:
                        print("🔴 Opção inválida. Tente novamente.")
                except ValueError:
                    print("🔴 Entrada inválida. Por favor, insira um número.")
        else:
            # Caso nenhum dado seja encontrado na tabela
            print(f"🔴 Nenhum dado encontrado na tabela {tabela}.")
            return None
    except Exception as e:
        # Captura e exibe mensagens de erro durante a execução
        print(f"\n🔴 Erro ao listar opções na tabela {tabela}: {e}")
        return None
    finally:
        fechar_conexao(conexao)  # Fecha a conexão com o Banco de Dados

# Valida números positivos
def validar_numero_positivo(valor: str, nome_campo: str) -> float:
    while True:
        try:
            numero = float(valor)  # Tenta converter o valor para float
            if numero <= 0:
                raise ValueError(f"O campo '{nome_campo}' deve ser um número positivo.")  # Valida se é positivo
            return numero  # Retorna o número validado
        except ValueError as e:
            # Exibe erro e solicita nova entrada do usuário
            print(f"\n🔴 {e}")
            valor = input(f"Insira novamente o campo '{nome_campo}': ")

# Valida strings não vazias
def validar_string_nao_vazia(valor: str, nome_campo: str) -> str:
    while True:
        if not valor.strip():  # Verifica se a string está vazia ou contém apenas espaços
            print(f"🔴 O campo '{nome_campo}' não pode estar vazio.")
            valor = input(f"Insira novamente o campo '{nome_campo}': ")  # Solicita uma nova entrada
        else:
            return valor.strip()  # Retorna a string validada sem espaços em branco

# Insere um novo projeto no Banco de Dados
def inserir_projeto() -> None:
    try:
        limpar_terminal()  # Limpa o terminal antes de exibir o menu
        print("\n=== Cadastrando um novo projeto ===")
        conexao = conectarBD()  # Estabelece conexão com o banco
        if not conexao:
            return
        cursor = conexao.cursor()

        # Coleta e valida as informações do projeto
        descricao = validar_string_nao_vazia(input("Descrição do projeto: "), "Descrição")
        custo = validar_numero_positivo(input("Custo do projeto: "), "Custo")

        # Menu para selecionar o status do projeto
        print("\n=== Escolha o status do projeto ===")
        print("1. Em andamento")
        print("2. Concluído")
        while True:
            status_opcao = input("Escolha uma opção (1-2): ").strip()
            if status_opcao == "1":
                status = "Em andamento"  # Define o status como "Em andamento"
                break
            elif status_opcao == "2":
                status = "Concluído"  # Define o status como "Concluído"
                break
            else:
                print("🔴 Opção inválida. Tente novamente.")  # Solicita uma escolha válida

        # Lista opções para tipo de fonte e região
        id_tipo_fonte = listar_opcoes("TBL_TIPO_FONTES", "ID_TIPO_FONTE", "NOME")
        id_regiao = listar_opcoes("TBL_REGIOES_SUSTENTAVEIS", "ID_REGIAO", "NOME")

        # Cancela a operação se não conseguir selecionar as opções
        if id_tipo_fonte is None or id_regiao is None:
            print("🔴 Operação cancelada devido a falha na seleção de dados.")
            return

        # Insere os dados do projeto no banco de dados
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
        conexao.commit()  # Confirma as alterações no banco
        print("\n🟢 Projeto inserido com sucesso!")
    except Exception as e:
        # Mensagem de erro em caso de falha
        print(f"\n🔴 Erro ao inserir projeto: {e}")
    finally:
        fechar_conexao(conexao)  # Fecha a conexão com o Banco de Dados
        input("\nPressione Enter para continuar...")
        
# Atualiza um projeto existente no Banco de Dados
def atualizar_projeto() -> None:
    try:
        limpar_terminal()  # Limpa o terminal para exibição organizada
        print("\n=== Atualizando um projeto ===")
        conexao = conectarBD()  # Conecta ao banco de dados
        if not conexao:
            return
        cursor = conexao.cursor()

        # Solicita o ID do projeto a ser atualizado
        id_projeto = validar_numero_positivo(input("ID do projeto a ser atualizado: "), "ID do Projeto")

        # Consulta informações atuais do projeto
        consulta = """
            SELECT 
                p.ID_PROJETO, 
                p.DESCRICAO, 
                p.CUSTO, 
                p.STATUS, 
                p.ID_TIPO_FONTE, 
                tf.NOME AS TIPO_FONTE, 
                p.ID_REGIAO, 
                r.NOME AS REGIAO
            FROM RM556310.TBL_PROJETOS_SUSTENTAVEIS p
            JOIN TBL_TIPO_FONTES tf ON p.ID_TIPO_FONTE = tf.ID_TIPO_FONTE
            JOIN TBL_REGIOES_SUSTENTAVEIS r ON p.ID_REGIAO = r.ID_REGIAO
            WHERE p.ID_PROJETO = :id_projeto
        """
        cursor.execute(consulta, {"id_projeto": id_projeto})
        resultado = cursor.fetchone()

        if not resultado:  # Verifica se o projeto foi encontrado
            print("\n🔴 Projeto com o ID informado não encontrado.")
            input("\nPressione Enter para continuar...")
            return

        # Carrega os dados do projeto
        projeto_atual = {
            "ID_PROJETO": resultado[0],
            "DESCRICAO": resultado[1],
            "CUSTO": resultado[2],
            "STATUS": resultado[3],
            "ID_TIPO_FONTE": resultado[4],
            "TIPO_FONTE": resultado[5],
            "ID_REGIAO": resultado[6],
            "REGIAO": resultado[7],
        }

        # Clona os dados originais para verificar alterações posteriormente
        projeto_inicial = projeto_atual.copy()

        while True:
            limpar_terminal()  # Limpa o terminal para exibição organizada
            print("\n=== Informações atuais do projeto ===")
            # Exibe os dados do projeto no formato correto
            print(f"ID: {projeto_atual['ID_PROJETO']}")
            print(f"Descrição: {projeto_atual['DESCRICAO']}")
            print(f"Custo: R${projeto_atual['CUSTO']:,.2f}")
            print(f"Status: {projeto_atual['STATUS']}")
            print(f"Tipo de Fonte ID: {projeto_atual['ID_TIPO_FONTE']} ({projeto_atual['TIPO_FONTE']})")
            print(f"Região ID: {projeto_atual['ID_REGIAO']} ({projeto_atual['REGIAO']})")

            # Menu de opções para atualizar os campos
            print("\n=== Escolha o campo que deseja modificar ===")
            print("1. Descrição")
            print("2. Custo")
            print("3. Status")
            print("4. Tipo de Fonte")
            print("5. Região")
            print("6. Voltar ao menu principal")

            opcao = input("Escolha uma opção (1-6): ").strip()

            if opcao == "1":
                # Atualizar descrição
                descricao = input("\nNova descrição do projeto: ").strip()
                if descricao:
                    projeto_atual["DESCRICAO"] = descricao
                    print("\n🟢 Descrição atualizada com sucesso!")
                else:
                    print("\n🔴 A descrição não pode ser vazia.")

            elif opcao == "2":
                # Atualizar custo
                custo = input("\nNovo custo do projeto: ").strip()
                if custo:
                    try:
                        projeto_atual["CUSTO"] = float(custo)
                        print("\n🟢 Custo atualizado com sucesso!")
                    except ValueError:
                        print("\n🔴 O custo deve ser um número válido.")
                else:
                    print("\n🔴 O custo não pode ser vazio.")

            elif opcao == "3":
                # Atualizar status
                print("\n=== Escolha o novo status do projeto ===")
                print("1. Em andamento")
                print("2. Concluído")
                while True:
                    status_opcao = input("Escolha uma opção (1-2): ").strip()
                    if status_opcao == "1":
                        projeto_atual["STATUS"] = "Em andamento"
                        print("\n🟢 Status atualizado para 'Em andamento'.")
                        break
                    elif status_opcao == "2":
                        projeto_atual["STATUS"] = "Concluído"
                        print("\n🟢 Status atualizado para 'Concluído'.")
                        break
                    else:
                        print("🔴 Opção inválida. Tente novamente.")

            elif opcao == "4":
                # Atualizar tipo de fonte
                print("\n=== Escolha o novo tipo de fonte ===")
                id_tipo_fonte = listar_opcoes("TBL_TIPO_FONTES", "ID_TIPO_FONTE", "NOME")
                if id_tipo_fonte is not None:
                    projeto_atual["ID_TIPO_FONTE"] = id_tipo_fonte
                    consulta_tipo = """
                        SELECT NOME FROM TBL_TIPO_FONTES WHERE ID_TIPO_FONTE = :id_tipo_fonte
                    """
                    cursor.execute(consulta_tipo, {"id_tipo_fonte": id_tipo_fonte})
                    novo_tipo = cursor.fetchone()
                    projeto_atual["TIPO_FONTE"] = novo_tipo[0] if novo_tipo else "Desconhecido"
                    print("\n🟢 Tipo de fonte atualizado com sucesso!")

            elif opcao == "5":
                # Atualizar região
                print("\n=== Escolha a nova região do projeto ===")
                id_regiao = listar_opcoes("TBL_REGIOES_SUSTENTAVEIS", "ID_REGIAO", "NOME")
                if id_regiao is not None:
                    projeto_atual["ID_REGIAO"] = id_regiao
                    consulta_regiao = """
                        SELECT NOME FROM TBL_REGIOES_SUSTENTAVEIS WHERE ID_REGIAO = :id_regiao
                    """
                    cursor.execute(consulta_regiao, {"id_regiao": id_regiao})
                    nova_regiao = cursor.fetchone()
                    projeto_atual["REGIAO"] = nova_regiao[0] if nova_regiao else "Desconhecida"
                    print("\n🟢 Região atualizada com sucesso!")

            elif opcao == "6":
                # Salva as alterações no banco de dados e encerra
                if projeto_atual != projeto_inicial:
                    query = """
                        UPDATE TBL_PROJETOS_SUSTENTAVEIS
                        SET DESCRICAO = :descricao, CUSTO = :custo, STATUS = :status, 
                            ID_TIPO_FONTE = :id_tipo_fonte, ID_REGIAO = :id_regiao
                        WHERE ID_PROJETO = :id_projeto
                    """
                    cursor.execute(
                        query,
                        {
                            "descricao": projeto_atual["DESCRICAO"],
                            "custo": projeto_atual["CUSTO"],
                            "status": projeto_atual["STATUS"],
                            "id_tipo_fonte": projeto_atual["ID_TIPO_FONTE"],
                            "id_regiao": projeto_atual["ID_REGIAO"],
                            "id_projeto": projeto_atual["ID_PROJETO"],
                        },
                    )
                    conexao.commit()
                    print("\n🟢 Projeto atualizado com sucesso!")
                else:
                    print("\n🔵 Nenhuma alteração foi realizada.")
                input("\nPressione Enter para continuar...")
                break
            else:
                print("\n🔴 Opção inválida. Tente novamente.")

            # Pergunta se o usuário deseja alterar mais campos
            alterar_mais = input("\nDeseja modificar mais algum campo? (s/n): ").strip().lower()
            if alterar_mais != "s":
                if projeto_atual != projeto_inicial:
                    # Salva alterações
                    query = """
                        UPDATE TBL_PROJETOS_SUSTENTAVEIS
                        SET DESCRICAO = :descricao, CUSTO = :custo, STATUS = :status, 
                            ID_TIPO_FONTE = :id_tipo_fonte, ID_REGIAO = :id_regiao
                        WHERE ID_PROJETO = :id_projeto
                    """
                    cursor.execute(
                        query,
                        {
                            "descricao": projeto_atual["DESCRICAO"],
                            "custo": projeto_atual["CUSTO"],
                            "status": projeto_atual["STATUS"],
                            "id_tipo_fonte": projeto_atual["ID_TIPO_FONTE"],
                            "id_regiao": projeto_atual["ID_REGIAO"],
                            "id_projeto": projeto_atual["ID_PROJETO"],
                        },
                    )
                    conexao.commit()
                    print("\n🟢 Todas as alterações foram salvas com sucesso!")
                else:
                    print("\n🔵 Nenhuma alteração foi realizada.")
                input("\nPressione Enter para continuar...")
                break

    except Exception as e:
        # Mensagem de erro em caso de falha
        print(f"\n🔴 Erro ao atualizar projeto: {e}")
        input("\nPressione Enter para continuar...")
    finally:
        fechar_conexao(conexao)  # Fecha a conexão com o Banco de Dados

# Exclui um projeto existente do Banco de Dados
def excluir_projeto() -> None:
    try:
        limpar_terminal()  # Limpa o terminal para exibição organizada
        print("\n=== Excluindo um projeto ===")
        conexao = conectarBD()  # Estabelece conexão com o banco de dados
        if not conexao:
            return
        cursor = conexao.cursor()

        # Solicita o ID do projeto a ser excluído
        id_projeto = validar_numero_positivo(input("ID do projeto a ser excluído: "), "ID do Projeto")

        # Consulta as informações do projeto para confirmar exclusão
        consulta = "SELECT DESCRICAO, CUSTO, STATUS FROM TBL_PROJETOS_SUSTENTAVEIS WHERE ID_PROJETO = :id_projeto"
        cursor.execute(consulta, {"id_projeto": id_projeto})
        projeto = cursor.fetchone()

        if not projeto:  # Verifica se o projeto foi encontrado
            print("\n🔴 Projeto com o ID informado não encontrado.")
            return

        # Exibe os detalhes do projeto antes de confirmar exclusão
        print("\n=== Informações do Projeto ===")
        print(f"Descrição: {projeto[0]}")
        print(f"Custo: R${projeto[1]:,.2f}")
        print(f"Status: {projeto[2]}")
        print("\nTem certeza que deseja excluir este projeto? (sim/não)")
        
        confirmacao = input("Digite sua escolha: ").strip().lower()
        if confirmacao != "sim":
            print("\n🔴 Exclusão cancelada pelo usuário.")  # Cancela exclusão se não confirmado
            return

        # Realiza a exclusão do projeto no banco de dados
        query = "DELETE FROM TBL_PROJETOS_SUSTENTAVEIS WHERE ID_PROJETO = :id_projeto"
        cursor.execute(query, {"id_projeto": id_projeto})
        conexao.commit()  # Salva a exclusão
        print("\n🟢 Projeto excluído com sucesso!")
    except Exception as e:
        # Exibe mensagem de erro em caso de falha
        print(f"\n🔴 Erro ao excluir projeto: {e}")
    finally:
        fechar_conexao(conexao)  # Fecha a conexão com o Banco de Dados
        input("\nPressione Enter para continuar...")

# Consulta os projetos existentes no Banco de Dados
def consultar_projetos(export: bool = False) -> list:
    try:
        if export:
            print("\n=== Selecione os projetos que deseja exportar ===")
        else:
            limpar_terminal()  # Limpa o terminal para exibição organizada
            print("\n=== Consultando projetos ===")
        
        # Menu de filtro para a consulta de projetos
        print("1. Todos os projetos")
        print("2. Apenas os projetos em andamento")
        print("3. Apenas os projetos concluídos")

        escolha = input("Escolha uma opção (1-3): ").strip()
        if escolha not in ["1", "2", "3"]:
            print("\n🔴 Opção inválida.")  # Retorna se a escolha for inválida
            return []

        # Monta a consulta SQL com base na escolha do usuário
        consulta = """
            SELECT 
                p.ID_PROJETO, 
                p.DESCRICAO, 
                p.CUSTO, 
                p.STATUS, 
                tf.ID_TIPO_FONTE, 
                tf.NOME AS TIPO_FONTE, 
                r.ID_REGIAO, 
                r.NOME AS REGIAO
            FROM RM556310.TBL_PROJETOS_SUSTENTAVEIS p
            JOIN TBL_TIPO_FONTES tf ON p.ID_TIPO_FONTE = tf.ID_TIPO_FONTE
            JOIN TBL_REGIOES_SUSTENTAVEIS r ON p.ID_REGIAO = r.ID_REGIAO
        """
        if escolha == "2":
            consulta += " WHERE p.STATUS = 'Em andamento'"  # Filtra projetos em andamento
        elif escolha == "3":
            consulta += " WHERE p.STATUS = 'Concluído'"  # Filtra projetos concluídos

        conexao = conectarBD()  # Conecta ao banco de dados
        if not conexao:
            return []
        cursor = conexao.cursor()
        cursor.execute(consulta)

        resultados = cursor.fetchall()  # Armazena os resultados da consulta
        if not resultados:
            print("\n🔴 Nenhum projeto encontrado.")  # Exibe mensagem se não houver resultados
        else:
            if export:
                print("\n=== Projetos que serão exportados ===")
            # Exibe os resultados da consulta
            for projeto in resultados:
                print(
                    f"\nID: {projeto[0]} | Descrição: {projeto[1]} | "
                    f"Custo: R${projeto[2]:,.2f} | Status: {projeto[3]} | "
                    f"Tipo de Fonte ID: {projeto[4]} ({projeto[5]}) | Região ID: {projeto[6]} ({projeto[7]})"
                )

        if not export:
            input("\nPressione Enter para continuar...")  # Pausa para visualização dos resultados

        return resultados  # Retorna a lista de resultados
    finally:
        fechar_conexao(conexao)  # Fecha a conexão com o Banco de Dados

# Exporta projetos selecionados para um arquivo JSON
def exportar_json(dados: list, nome_arquivo: str = None) -> None:
    try:
        if not nome_arquivo:
            hoje = datetime.now().strftime("%Y-%m-%d")  # Gera um nome padrão para o arquivo
            nome_arquivo = f"projetos_{hoje}.json"

        # Transforma os dados em formato JSON
        if isinstance(dados, list) and all(isinstance(item, (list, tuple)) for item in dados):
            dados = [
                {
                    "ID_PROJETO": item[0],
                    "DESCRICAO": item[1],
                    "CUSTO": item[2],
                    "STATUS": item[3],
                    "ID_TIPO_FONTE": item[4],  # Apenas o ID
                    "ID_REGIAO": item[6]  # Apenas o ID
                }
                for item in dados
            ]

        # Escreve os dados em um arquivo JSON
        with open(nome_arquivo, "w", encoding="utf-8") as arquivo:
            json.dump(dados, arquivo, indent=4, ensure_ascii=False)
        print(f"\n🟢 Dados exportados para o arquivo: {nome_arquivo}")
    except Exception as e:
        # Exibe mensagem de erro caso a exportação falhe
        print(f"\n🔴 Erro ao exportar para JSON: {e}")
    input("\nPressione Enter para continuar...")  # Pausa para visualização da mensagem

# Exporta projetos selecionados para um arquivo Excel (.xlsx)
def exportar_DataFrame(dados: list, nome_arquivo: str = None) -> None:
    try:
        if not nome_arquivo:
            hoje = datetime.now().strftime("%Y-%m-%d")  # Gera um nome padrão para o arquivo
            nome_arquivo = f"projetos_{hoje}.xlsx"

        # Converte os dados para um DataFrame do pandas
        df = pd.DataFrame(
            [
                {
                    "ID_PROJETO": item[0],
                    "DESCRICAO": item[1],
                    "CUSTO": item[2],
                    "STATUS": item[3],
                    "ID_TIPO_FONTE": item[4],  # Apenas o ID
                    "ID_REGIAO": item[6]  # Apenas o ID
                }
                for item in dados
            ]
        )

        # Exporta o DataFrame para um arquivo Excel
        df.to_excel(nome_arquivo, index=False, engine="openpyxl")
        print(f"\n🟢 Dados exportados para o arquivo: {nome_arquivo}")
    except ModuleNotFoundError:
        # Mensagem caso a biblioteca necessária não esteja instalada
        print("\n🔴 O módulo 'openpyxl' não está instalado. Por favor, instale-o usando 'pip install openpyxl'.")
    except Exception as e:
        # Exibe mensagem de erro caso a exportação falhe
        print(f"\n🔴 Erro ao exportar para Excel: {e}")
    input("\nPressione Enter para continuar...")  # Pausa para visualização

# Exibe o menu principal
def exibir_menu() -> None:
    limpar_terminal()  # Limpa o terminal antes de exibir o menu
    print("\n=== MENU PRINCIPAL ===")
    # Opções disponíveis no sistema
    print("1. Cadastrar novo Projeto")
    print("2. Atualizar Projeto Existente")
    print("3. Excluir Projeto pelo ID")
    print("4. Consultar Projetos pelo status")
    print("5. Exportar Projetos para JSON ou DataFrame")
    print("6. Sair")

# Função principal que controla o fluxo do programa
def main() -> None:
    while True:
        exibir_menu()  # Exibe o menu principal
        opcao = input("Escolha uma opção: ")  # Solicita a escolha do usuário
        if opcao == "1":
            inserir_projeto()  # Chama a função para cadastrar um projeto
        elif opcao == "2":
            atualizar_projeto()  # Chama a função para atualizar um projeto
        elif opcao == "3":
            excluir_projeto()  # Chama a função para excluir um projeto
        elif opcao == "4":
            consultar_projetos()  # Chama a função para consultar projetos
        elif opcao == "5":
            # Realiza a exportação de projetos
            projects = consultar_projetos(export=True)
            if projects:
                while True:
                    print("\n=== Exportar Dados ===")
                    print("1. Exportar para JSON")
                    print("2. Exportar para Excel")
                    export_opcao = input("Escolha uma opção (1-2): ").strip()
                    if export_opcao == "1":
                        exportar_json(projects)  # Exporta para JSON
                        break
                    elif export_opcao == "2":
                        exportar_DataFrame(projects)  # Exporta para Excel
                        break
                    else:
                        print("🔴 Opção inválida. Tente novamente.")
            else:
                print("🔴 Nenhum dado disponível para exportação.")
        elif opcao == "6":
            # Finaliza o sistema
            print("\n🟢 Saindo do sistema...")
            break
        else:
            # Mensagem para opções inválidas
            print("\n🔴 Opção inválida. Tente novamente.")
            input("\nPressione Enter para continuar...")

# Executa a função principal
main()