# --- gerar_senhas.py ---
import streamlit_authenticator as stauth

# Coloque aqui as senhas que você quer usar para cada usuário
senhas_em_texto_plano = ['andrade1994', 'thomas123', 'thilia123', 'neto123', 'admin123'] 

hashed_passwords = stauth.Hasher(senhas_em_texto_plano).generate()
print("Copie as senhas criptografadas abaixo e cole no seu arquivo config.yaml:")
print(hashed_passwords)