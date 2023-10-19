# Script de Automação Web com Selenium
Este script em Python utiliza a biblioteca Selenium para automatizar interações com um aplicativo da web. Ele foi projetado para um caso de uso específico no site "magalu.brainlaw.com.br".

## Pré-requisitos
Antes de executar o script, certifique-se de ter o seguinte configurado:

- Python instalado no seu sistema.
- Chrome WebDriver instalado e compatível com o seu navegador Chrome. Você pode baixá-lo [aqui](https://googlechromelabs.github.io/chrome-for-testing/).
- Bibliotecas Python necessárias instaladas (instale com pip):
- Selenium
- Pandas
- NumPy
- PyAutoGUI
- dotenv

## Instalação
Clone ou faça o download do repositório para o seu computador local.

Instale as bibliotecas Python necessárias usando o seguinte comando:

```
pip install selenium pandas numpy pyautogui python-dotenv
```

Baixe o Chrome WebDriver e coloque-o no mesmo diretório do script.

Crie um arquivo .env no diretório do projeto com o seguinte conteúdo:

```
LOGIN_BL=seu_login_magalu
PASSWORD_BL=sua_senha_magalu
```

Substitua seu_login_magalu e sua_senha_magalu pelas suas credenciais de login reais.

## Uso

Execute o script:

python nome_do_seu_script.py

O script automatizará as interações com o site "magalu.brainlaw.com.br" com base nos dados fornecidos no arquivo "Eventos.xlsx".
