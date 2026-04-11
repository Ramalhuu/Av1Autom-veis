# Av1Automóveis — GitHub Pages + Gerenciador Local

## 📁 Estrutura do Projeto

```
av1automoveis/
│
├── 📄 index.html              ← Site completo (5 páginas)
├── 📄 carros.json             ← Banco de dados (lido pelo site)
│
├── images/                    ← Fotos dos carros
│   ├── carro_1.jpg
│   └── ...
│
├── gerenciador/
│   └── 📄 gerenciador.py     ← App de gestão do estoque
│
└── data/
    └── 📊 carros.xlsx        ← Banco de dados original (Excel)
```

---

## 🚀 PASSO A PASSO: Publicar no GitHub Pages

### 1. Criar conta no GitHub (se não tiver)
Acesse https://github.com e crie uma conta gratuita.

### 2. Criar repositório
1. Clique em **"New repository"**
2. Nome: `av1automoveis` (ou qualquer nome)
3. Deixe como **Public**
4. Clique em **"Create repository"**

### 3. Fazer upload dos arquivos
Na tela do repositório recém-criado:
1. Clique em **"uploading an existing file"**
2. Arraste TODOS os arquivos desta pasta:
   - `index.html`
   - `carros.json`
   - Pasta `images/` (com as fotos)
3. Clique em **"Commit changes"**

> ⚠️ A pasta `gerenciador/` e `data/` NÃO precisam ir para o GitHub.
> Fique com elas só no seu computador.

### 4. Ativar o GitHub Pages
1. No repositório, clique em **Settings** (engrenagem)
2. No menu lateral, clique em **Pages**
3. Em "Branch", selecione **main** e clique **Save**
4. Aguarde ~1 minuto

Seu site estará em:
```
https://SEU_USUARIO.github.io/av1automoveis/
```

---

## 🔄 FLUXO DE TRABALHO: Adicionar ou Editar Carros

```
1. Abra o gerenciador.py no seu computador
2. Adicione/edite/remova veículos
3. O gerenciador salva no carros.xlsx
   e gera carros.json automaticamente
4. Acesse o GitHub no navegador
5. Delete o carros.json antigo
6. Faça upload do carros.json novo
7. (Se adicionou foto nova) Upload da imagem em images/
8. Site atualizado em ~1 minuto!
```

### Opção mais rápida: usar GitHub Desktop
1. Instale: https://desktop.github.com
2. Clone seu repositório
3. Após exportar o JSON pelo gerenciador, arraste o `carros.json`
   para a pasta clonada
4. No GitHub Desktop: escreva uma mensagem e clique **Commit** → **Push**

---

## ⚙️ Configurar o WhatsApp

Abra `index.html` e na linha que diz:
```javascript
const WPP = '5511999999999';
```
Troque pelo número real (só números, com DDI 55 + DDD + número):
```javascript
const WPP = '5511987654321';
```

---

## 🖥️ Pré-requisitos para o Gerenciador

Python 3.8+ instalado. Rode uma vez:
```bash
pip install openpyxl Pillow
```

Para abrir o gerenciador:
```bash
python gerenciador/gerenciador.py
```
Ou use o arquivo `ABRIR_GERENCIADOR.bat` (Windows, duplo clique).

---

## 📌 Dicas Importantes

| Situação | O que fazer |
|----------|-------------|
| Adicionei um carro | Exportar JSON → upload no GitHub |
| Marquei como vendido | Exportar JSON → upload no GitHub |
| Adicionei foto | Upload do `carros.json` + arquivo da imagem em `images/` |
| Site não atualiza | Espere 1-2 min ou force F5 no navegador |
| Imagem não aparece | Confirme que o arquivo está em `images/` no GitHub |

---

## 🔧 Solução de Problemas

**Site mostra dados antigos?**
O GitHub Pages tem cache. Adicione `?v=2` na URL para forçar atualização.

**Imagem não aparece no site?**
Confirme que o `imagem_path` no JSON é `images/nome_do_arquivo.jpg`
e que o arquivo existe na pasta `images/` do repositório.

**Erro "Failed to fetch carros.json"?**
O site usa dados de demonstração até encontrar o arquivo.
Certifique-se que `carros.json` está na raiz do repositório.
