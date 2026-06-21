# ICMS/GO — Plataforma de Análise Fiscal

Aplicação web Streamlit para análise tributária de ICMS para empresas varejistas no Estado de Goiás.

## Funcionalidades

- **Análise de Entradas** — identifica divergências de ICMS em notas fiscais de entrada (DIV-1 a DIV-5), gera Excel + Word
- **Análise de Saídas** — identifica divergências de ICMS em notas fiscais de saída (DIV-1 a DIV-7), gera Excel + Word
- **Apuração Mensal** — calcula débito, crédito, DIFAL, PROTEGE/GO e saldo a recolher por filial, gera planilha com 3 abas

## Estrutura de arquivos

Todos os arquivos devem estar na **mesma pasta**:

```
icms-go-app/
├── app.py                    ← aplicação Streamlit (este projeto)
├── requirements.txt
├── README.md
├── .streamlit/
│   └── config.toml
│
├── analisar_entradas.py      ← copiar do projeto original
├── analisar_saidas.py        ← copiar do projeto original
├── apuracao_3abas.py         ← copiar do projeto original
└── combinar_xlsx.py          ← copiar de apuracao-icms-go/scripts/
```

## Deploy no Streamlit Community Cloud

### Pré-requisitos
- Conta gratuita em [streamlit.io](https://streamlit.io)
- Conta no GitHub

### Passo 1 — Preparar os arquivos

1. Copie os scripts de análise para a pasta `icms-go-app/`:
   ```
   analisar_entradas.py
   analisar_saidas.py
   apuracao_3abas.py
   combinar_xlsx.py   (de apuracao-icms-go/scripts/)
   ```

2. Verifique que a estrutura de arquivos está correta (veja acima)

### Passo 2 — Criar repositório GitHub

1. Acesse [github.com](https://github.com) e crie um **novo repositório privado** (ex: `icms-go-app`)
2. Faça upload de todos os arquivos da pasta `icms-go-app/` para o repositório
   - `app.py`
   - `requirements.txt`
   - `analisar_entradas.py`
   - `analisar_saidas.py`
   - `apuracao_3abas.py`
   - `combinar_xlsx.py`
   - `.streamlit/config.toml`

### Passo 3 — Conectar ao Streamlit Community Cloud

1. Acesse [share.streamlit.io](https://share.streamlit.io)
2. Clique em **"New app"**
3. Selecione seu repositório GitHub
4. Configure:
   - **Branch:** `main`
   - **Main file path:** `app.py`
5. Clique em **"Advanced settings..."**

### Passo 4 — Configurar a senha (Secrets)

Na tela de configurações avançadas, adicione o seguinte em **Secrets**:

```toml
senha = "sua_senha_aqui"
```

> ⚠️ Substitua `"sua_senha_aqui"` pela senha que deseja usar.
> Se não configurar o secret, a senha padrão será `icms2026`.

### Passo 5 — Deploy

1. Clique em **"Deploy!"**
2. Aguarde o build (2-5 minutos na primeira vez)
3. A aplicação estará disponível em `https://seu-usuario-icms-go-app.streamlit.app`

## Uso da aplicação

1. Acesse a URL da aplicação
2. Digite a senha configurada
3. Selecione a aba desejada
4. Faça upload dos arquivos XLS/CSV de movimentação fiscal
5. Clique em "Processar"
6. Baixe os relatórios gerados (Excel e/ou Word)

## Formatos de arquivo aceitos

| Extensão | Descrição |
|----------|-----------|
| `.xls`   | Excel legado — exportação padrão de ERPs |
| `.xlsx`  | Excel moderno |
| `.csv`   | CSV com separador `;` ou `,`, encoding Latin-1 |

## Observações técnicas

- Os arquivos são processados em memória (pasta temporária) e **não são armazenados** no servidor
- Limite de upload: 200 MB por arquivo (configurável em `.streamlit/config.toml`)
- Timeout de processamento: 5 minutos para a apuração mensal
- A aplicação é stateless — cada sessão é independente
- Para arquivos muito grandes (>50k linhas), o processamento pode levar alguns minutos

## Suporte

Em caso de dúvidas ou erros, verifique:
1. Se todos os scripts estão na mesma pasta que `app.py`
2. Se os arquivos de movimentação estão no formato correto (colunas esperadas pelo ERP)
3. O log de processamento exibido na interface após o processamento
