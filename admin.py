# =====================================================================
# RWA TECNOLOGIA OPERACIONAL — PAINEL ADMIN (linha de comando)
# Uso: python admin.py
# =====================================================================

import sys
import secrets
import database
import auth

database.inicializar_banco()


def gerar_chave_aes() -> str:
    return secrets.token_hex(32).upper()


def menu():
    print("\n" + "═" * 52)
    print("  RWA TECNOLOGIA OPERACIONAL — ADMIN")
    print("═" * 52)
    print("  1. Cadastrar empresa")
    print("  2. Listar empresas")
    print("  3. Listar máquinas de uma empresa")
    print("  4. Remover máquina  (libera nova instalação)")
    print("  5. Registrar máquina (adicionar fingerprint)")
    print("  6. Renovar licença")
    print("  7. Bloquear empresa")
    print("  0. Sair")
    print("═" * 52)
    return input("  Opção: ").strip()


def main():
    while True:
        op = menu()

        if op == "1":
            print("\n── CADASTRAR EMPRESA ──")
            nome       = input("Nome da empresa         : ").strip()
            documento  = input("CNPJ/CPF                : ").strip()
            email      = input("Email de acesso         : ").strip()
            telefone   = input("Telefone                : ").strip()
            vencimento = input("Vencimento (YYYY-MM-DD) : ").strip()
            chave      = gerar_chave_aes()
            print(f"\n  Chave AES gerada: {chave}")
            ok = input("  Confirmar? (s/n): ").strip().lower()
            if ok == "s":
                database.cadastrar_empresa(nome, email, None, vencimento, chave, documento, telefone)
                print("  ✓ Empresa cadastrada.")

        elif op == "2":
            print("\n── EMPRESAS CADASTRADAS ──")
            database.listar_empresas()

        elif op == "3":
            print("\n── MÁQUINAS DA EMPRESA ──")
            eid = int(input("ID da empresa: ").strip())
            database.listar_maquinas(eid)

        elif op == "4":
            print("\n── REMOVER MÁQUINA ──")
            eid = int(input("ID da empresa : ").strip())
            fp  = input("Fingerprint   : ").strip()
            database.remover_maquina(eid, fp)

        elif op == "5":
            print("\n── REGISTRAR MÁQUINA ──")
            email = input("Email da empresa  : ").strip()
            fp    = input("Fingerprint       : ").strip()
            emp   = database.buscar_empresa(email)
            if not emp:
                print("  ✗ Empresa não encontrada.")
            else:
                database.registrar_maquina(emp["id"], fp)
                print("  ✓ Máquina registrada.")

        elif op == "6":
            print("\n── RENOVAR LICENÇA ──")
            email = input("Email da empresa        : ").strip()
            venc  = input("Novo vencimento (YYYY-MM-DD): ").strip()
            database.renovar_licenca(email, venc)

        elif op == "7":
            print("\n── BLOQUEAR EMPRESA ──")
            email = input("Email da empresa: ").strip()
            database.bloquear_empresa(email)

        elif op == "0":
            print("\nSaindo...\n")
            sys.exit(0)

        else:
            print("  Opção inválida.")


if __name__ == "__main__":
    main()
