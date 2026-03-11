# Plano de Curso – UFAC

Formulário web para geração de Planos de Curso da UFAC. O arquivo `.docx` é gerado diretamente no navegador, sem necessidade de servidor.

## Como publicar no GitHub Pages

### 1. Criar repositório no GitHub
- Acesse [github.com](https://github.com) e faça login
- Clique em **New repository**
- Dê um nome, ex: `plano-curso-ufac`
- Deixe como **Public**
- Clique em **Create repository**

### 2. Fazer upload dos arquivos
- Na página do repositório, clique em **Add file → Upload files**
- Arraste os três arquivos:
  - `index.html`
  - `disciplinas.json`
  - `logo.png`
- Clique em **Commit changes**

### 3. Ativar o GitHub Pages
- Vá em **Settings** (aba no topo do repositório)
- No menu lateral, clique em **Pages**
- Em **Source**, selecione a branch `main` e pasta `/ (root)`
- Clique em **Save**

### 4. Acessar o site
Após alguns minutos, o site estará disponível em:
```
https://SEU-USUARIO.github.io/plano-curso-ufac
```

## Arquivos necessários
| Arquivo | Descrição |
|---|---|
| `index.html` | Formulário completo + geração do DOCX |
| `disciplinas.json` | Base de disciplinas do curso |
| `logo.png` | Logo da UFAC |

## Funcionalidades
- Preenchimento automático ao selecionar disciplina
- Geração do `.docx` 100% no navegador (sem servidor)
- Compatível com Microsoft Word e LibreOffice
