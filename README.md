# API de Conteúdos de IES

Uma API RESTful desenvolvida em Flask para fornecer acesso estruturado aos conteúdos de diferentes Instituições de Ensino Superior (IES) a partir de arquivos Excel.

## 📋 Funcionalidades

- ✅ Leitura de arquivos Excel com múltiplas abas (cada aba representa uma IES)
- ✅ Endpoints dinâmicos baseados nos nomes das IES
- ✅ Estrutura hierárquica dos conteúdos: Matéria → Tema → Subtema → Aula
- ✅ Interface web interativa com links clicáveis
- ✅ Suporte a cache para melhor performance
- ✅ Recarregamento de dados sem reiniciar o servidor
- ✅ Visualização em JSON e HTML

## 🚀 Como Executar

### Pré-requisitos

- Python 3.6 ou superior
- pip (gerenciador de pacotes do Python)

### Instalação

1. Clone o repositório:
```bash
git clone <url-do-repositorio>
cd <nome-do-repositorio>
