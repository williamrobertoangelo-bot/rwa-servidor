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
    print("  5. Renovar licença")
    print("  6. Bloquear empresa")
    print("  0. Sair")
    print("═" * 52)
    return input("  Opção: ").strip()


def main():
    while True:
        op = menu()

        if op == "1":
            print("\n── CADASTRAR EMPRESA ──")
            nome       = input("Nome da empresa         : ").strip()
            email      = input("Email de acesso         : ").strip()
            senha      = input("Senha                   : ").strip()
            vencimento = input("Vencimento (YYYY-MM-DD) : ").strip()
            chave      = gerar_chave_aes()
            senha_hash = auth.hash_senha(senha)
            print(f"\n  Chave AES gerada: {chave}")
            ok = input("  Confirmar? (s/n): ").strip().lower()
            if ok == "s":
                database.cadastrar_empresa(nome, email, senha_hash, vencimento, chave)
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
            print("\n── RENOVAR LICENÇA ──")
            email = input("Email da empresa        : ").strip()
            venc  = input("Novo vencimento (YYYY-MM-DD): ").strip()
            database.renovar_licenca(email, venc)

        elif op == "6":
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
