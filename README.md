# GFRT FX Exposure — Streamlit App

Dashboard de FX Exposure para o Grupo Fernando Ribas Taques.

## Estrutura de arquivos

```
gfrt_fx/
├── app.py                        ← App principal
├── requirements.txt              ← Dependências Python
├── .streamlit/
│   └── config.toml               ← Tema e configurações
└── GFRT_FX_Exposure_PowerBI.xlsx ← Base de dados (coloque aqui)
```

---

## Rodar localmente

```bash
# 1. Instalar dependências
pip install -r requirements.txt

# 2. Colocar o Excel na mesma pasta do app.py
cp /caminho/para/GFRT_FX_Exposure_PowerBI.xlsx .

# 3. Rodar
streamlit run app.py
```

Acessa em: http://localhost:8501

---

## Publicar no Streamlit Cloud (gratuito)

### Passo 1 — GitHub
1. Crie um repositório no GitHub (pode ser privado)
2. Faça upload de todos os arquivos desta pasta:
   - `app.py`
   - `requirements.txt`
   - `.streamlit/config.toml`
   - `GFRT_FX_Exposure_PowerBI.xlsx`

### Passo 2 — Streamlit Cloud
1. Acesse https://share.streamlit.io
2. Clique em **New app**
3. Conecte sua conta GitHub
4. Selecione o repositório e o branch
5. Em **Main file path**: `app.py`
6. Clique em **Deploy**

Pronto — o app fica disponível em um link público ou privado.

### Atualizar os dados
Basta substituir o arquivo `.xlsx` no repositório GitHub e dar push.
O Streamlit Cloud atualiza automaticamente em segundos.

---

## Páginas do dashboard

| Página | Conteúdo |
|---|---|
| Balanço FX | Ativo × passivo consolidado, waterfall, composição |
| Exposição Mensal | Fluxo mês a mês, FX acumulado, tabela SHORT/LONG |
| Soja / Comercialização | Produção, contratos, % vendido por fazenda |
| Insumos USD | Fornecedores, tipos, vencimentos, por fazenda |
| Dívida e Terras | Bancos, amortização por ano, parcelas Renascer |

---

## Atualização dos dados

O app lê diretamente do Excel. Para atualizar:
1. Atualize os dados nas planilhas de origem (25/26 e 26/27)
2. Regere o `GFRT_FX_Exposure_PowerBI.xlsx` com o script Python
3. Substitua o arquivo no repositório
