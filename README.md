# Instruções para uso do gerador de evidências

## Objetivo
Este projeto gera documentos Word a partir de uma planilha Excel e de um modelo Word.
O script lê a planilha, captura os campos de cada cenário e cria um arquivo `.docx` para cada linha com conteúdo na coluna `Title`.

---

## Estrutura de pastas esperada
O script foi configurado para usar estas pastas:

### 1. Pasta do projeto
Local onde está o script principal:
- `ProcessCreation.py`
- `instruction.md`

### 2. Pasta de entrada
Caminho esperado:
- `~/Downloads/Evidences/Input`

Nesta pasta, o usuário deve colocar:
- `BDD*.xlsx` → arquivo(s) Excel com os cenários
- `modelo.docx` → arquivo modelo que será usado como base para a criação dos documentos

> Se o arquivo `modelo.docx` não existir, o script cria um modelo padrão automaticamente.

### 3. Pasta de saída
Caminho esperado:
- `~/Downloads/Evidences/Output`

Nesta pasta o script salva os arquivos gerados.
Se a pasta não existir, ela é criada automaticamente.

---

## Como funciona
1. O script procura pela planilha Excel com padrão `BDD*.xlsx` dentro da pasta de entrada.
2. Lê a planilha usando a segunda linha como cabeçalho.
3. Para cada linha com valor na coluna `Title`, ele cria um documento Word.
4. O conteúdo dos campos `Given`, `When` e `Then` é adicionado ao documento.
5. O arquivo gerado recebe o mesmo nome da coluna `Title` (quando possível).

---

## Formato esperado da planilha
A planilha precisa conter, no mínimo, estas colunas:
- `ID`
- `Title`
- `Given`
- `When`
- `Then`

Exemplo de estrutura:

| ID | Title | Given | When | Then |
|---|---|---|---|---|
| CT01 | CT01 - Login com sucesso | DADO que... | QUANDO... | ENTÃO... |

> O script lê a coluna `Title` para definir o nome do arquivo final.

---

## Como executar
### Opção 1: via terminal
1. Abra o terminal na pasta do projeto.
2. Execute:
   ```bash
   python ProcessCreation.py
   ```

### Opção 2: usando ambiente virtual
Se o projeto tiver `.venv`, você pode executar:

```bash
source .venv/bin/activate
python ProcessCreation.py
```

---

## Resultado esperado
Após a execução, os arquivos `.docx` aparecerão em:
- `~/Downloads/Evidences/Output`

Cada arquivo será gerado com o conteúdo do cenário correspondente.

---

## Dicas importantes
- Não altere o nome da pasta `Input` ou `Output` caso queira que o script funcione sem ajustes.
- O arquivo Excel deve estar dentro da pasta `Input`.
- O arquivo `modelo.docx` deve estar na mesma pasta de entrada, ou o script criará um modelo padrão.
- Se o nome do arquivo final ficar com caracteres inválidos, o script tenta limpar o nome automaticamente.

---

## Exemplo de uso final
1. Crie a pasta:
   - `~/Downloads/Evidences/Input`
2. Cole o arquivo Excel com nome começando em `BDD`.
3. Cole o arquivo `modelo.docx` (se houver).
4. Execute o script.
5. Os documentos gerados aparecerão em:
   - `~/Downloads/Evidences/Output`
