
# Invesight Backend 🚀

Este repositório contém o backend da aplicação **Invesight**, uma plataforma de análise de ativos financeiros. O backend foi desenvolvido em **Python 3** utilizando o framework **FastAPI** para fornecer uma API RESTful eficiente, com integração ao banco de dados **MongoDB Atlas** e funcionalidades de **web scraping** para obter dados financeiros em tempo real.

## 🛠 Tecnologias

- **Python 3** 🐍
- **FastAPI** ⚡: Framework para a criação de APIs RESTful rápidas e eficientes.
- **MongoDB Atlas** 🗃️: Banco de dados em nuvem utilizado para armazenamento dos dados dos ativos e imagens das ações.
- **Web Scraping** 🌐: Utilizado para obter dados financeiros não disponíveis no banco de dados.
- **Algoritmos de Busca com Lógica Fuzzy** 🔍: Implementados para realizar buscas rápidas e dinâmicas de ativos financeiros.
- **Planilhas Excel e Google Sheets** 📊 (não em uso atualmente): Integração com planilhas para exibição de dados obtidos através de web scraping (em versões anteriores).

---

## 📋 Funcionalidades

- **API RESTful** 🔌: A API segue as normas de RESTful para fornecer endpoints de fácil uso e integração.
- **Conexão com MongoDB Atlas** 🗄️: O projeto se conecta ao MongoDB Atlas para armazenar e acessar dados dos ativos e imagens (logos das ações).
- **Web Scraping** 🧑‍💻: O backend utiliza técnicas de scraping para buscar dados financeiros de ativos que não estão no banco de dados.
- **Busca Dinâmica e Rápida** ⚡: Algoritmos de busca com lógica fuzzy são utilizados para realizar buscas rápidas de ativos no frontend.
- **EndPoints de Ativos do Tesouro Direto** 💰: Embora ainda não utilizados no frontend, a API contém endpoints para obtenção dos ativos do Tesouro Direto, que serão aprimorados quando implementados no futuro.

---

## 🗂 Estrutura de Diretórios

O projeto segue uma estrutura modular e organizada, com pastas e arquivos bem definidos para facilitar a manutenção e expansão. A estrutura básica é a seguinte:

```
invesight-backend/
│
├── AtivosAPI/
│   ├── entities/                # Definições dos modelos e entidades
│   │   └── actives.py
│   ├── functions/               # Funções de lógica de negócio
│   │   ├── database_functions/  # Funções para interagir com o banco de dados
│   │   │   ├── db_actives.py
│   │   │   ├── db_images.py
│   │   │   └── db_manipulation.py
│   │   ├── filter_functions/    # Funções de filtragem de dados
│   │   │   └── filter_functions.py
│   │   ├── request_functions/   # Funções de requisições
│   │   │   └── find_indices.py
│   │   └── utilities_functions/ # Funções utilitárias
│   │       └── utilities.py
│   ├── routes/                  # Endpoints da API
│   │   └── routes.py
│   └── web_scrapping/           # Scripts de web scraping
│       ├── b3_actives.py
│       └── treasury_bonds.py
│
├── PlanilhaExcel/
├── PlanilhaGoogle/
├── RasgagemDados/
├── .env
├── .gitignore
├── LICENSE
├── main.py
├── old_main.py
└── requirements.txt           # Dependências do projeto
```

---

## 🚀 Como Rodar o Projeto

### 🔧 Pré-requisitos

- Python 3.8 ou superior
- MongoDB Atlas configurado
- Dependências do projeto (listadas no `requirements.txt`)

### ⚙️ Instalação

1. Clone o repositório:

   ```bash
   git clone https://github.com/yourusername/invesight-backend.git
   cd invesight-backend
   ```

2. Instale as dependências:

   ```bash
   pip install -r requirements.txt
   ```

3. Configure as variáveis de ambiente (por exemplo, credenciais do MongoDB Atlas) no arquivo `.env`.

4. Para rodar a aplicação, execute:

   ```bash
   uvicorn app.main:app --reload
   ```

5. A API estará disponível em `http://localhost:8000`.

---

## 🤝 Como Contribuir

1. Faça um fork deste repositório.
2. Crie uma nova branch (`git checkout -b feature-nome-da-feature`).
3. Faça as modificações necessárias.
4. Envie as alterações (`git commit -am 'Adiciona nova funcionalidade'`).
5. Envie a branch para o repositório remoto (`git push origin feature-nome-da-feature`).
6. Crie um pull request.

---

## 📬 Contato

Criado por **Witor Oliveira**  
🔗 [LinkedIn](https://www.linkedin.com/in/witoroliveira/)  
📫 [Contato por e-mail](mailto:witoredson@gmail.com)

---

## 📜 Licença

Este projeto está licenciado sob a licença MIT. Veja o arquivo [MIT License](./LICENSE) para mais detalhes.
