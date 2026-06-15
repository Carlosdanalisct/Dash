# Instalação do Dashboard MMS

## 1. Instalar dependências

Abra o terminal na pasta do dashboard:

```text
outputs/dashboard_reclamacoes_app
```

Execute:

```bash
pip install -r requirements.txt
```

## 2. Iniciar o dashboard

Depois da instalação, use o arquivo:

```text
Iniciar Dashboard.bat
```

Ou execute manualmente:

```bash
python app.py --serve --host 127.0.0.1 --port 8787
```

## 3. Validação automática

Na inicialização, o dashboard verifica se as dependências obrigatórias estão instaladas.
Se alguma estiver ausente, ele mostra uma mensagem amigável indicando quais pacotes faltam e o comando:

```bash
pip install -r requirements.txt
```

