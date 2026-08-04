# XML Validator v1.0

Software desktop para comparar arquivos XML com um gabarito Excel.
**100% offline** - nenhum dado e enviado para a internet.

---

## Como usar (forma mais rapida)

1. Instale o **Python**: https://www.python.org/downloads/
   - Marque a opcao **"Add Python to PATH"** durante a instalacao
2. Clique duas vezes em **`instalar_e_rodar.bat`**
3. O programa abre automaticamente!

---

## Passo a passo no programa

| Passo | Acao |
|---|---|
| 1 | Clique em **Selecionar Excel** e carregue o gabarito |
| 2 | Clique em **Adicionar XMLs** e selecione os arquivos XML |
| 3 | Clique em **Validar XMLs** para iniciar a comparacao |
| 4 | Clique em **Exportar relatorio** para salvar o resultado |

---

## Formato do Excel (gabarito)

**Layout A - 2 colunas:**

| Campo | Valor Esperado |
|---|---|
| NomeCliente | Acme Corp |
| CNPJ | 00.000.000/0001-00 |

**Layout B - cabecalho horizontal:**

| NomeCliente | CNPJ | Cidade |
|---|---|---|
| Acme Corp | 00.000.000/0001-00 | Sao Paulo |

---

## Gerar EXE (para rodar sem Python)

1. Execute `gerar_exe.bat`
2. O arquivo `XMLValidator.exe` sera criado na pasta `dist/`
3. Copie o `.exe` para qualquer maquina Windows!

---

## Privacidade

Este software **nao envia nenhum dado para a internet**.
Tudo roda localmente dentro da sua rede.
