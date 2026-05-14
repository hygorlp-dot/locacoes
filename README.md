# 🔧 ArcD Locações — Sistema de Gestão de Equipamentos

App para cadastro e gestão de locação de equipamentos de obras (Betoneiras, Andaimes, Materlete, etc.), com controle por proprietário (Hygor e ArcD), alocação por obra e relatório mensal completo.

---

## 🚀 Deploy no Vercel + Supabase

### 1. Criar tabela no Supabase

Acesse [supabase.com](https://supabase.com), crie um projeto e rode este SQL no **SQL Editor**:

```sql
CREATE TABLE app_data (
  key   TEXT PRIMARY KEY,
  value TEXT NOT NULL
);

-- Permitir acesso público (anon key)
ALTER TABLE app_data ENABLE ROW LEVEL SECURITY;

CREATE POLICY "allow_all" ON app_data
  FOR ALL USING (true) WITH CHECK (true);
```

### 2. Obter as credenciais

No painel Supabase: **Settings → API**
- `Project URL` → será a `VITE_SUPABASE_URL`
- `anon public` key → será a `VITE_SUPABASE_ANON_KEY`

### 3. Deploy no Vercel

```bash
# Clone ou faça upload da pasta arced-locacoes no GitHub
# Depois no Vercel:
# 1. Import do repositório
# 2. Adicionar variáveis de ambiente:

VITE_SUPABASE_URL=https://xxxx.supabase.co
VITE_SUPABASE_ANON_KEY=eyJ...
```

Ou via CLI:
```bash
cd arced-locacoes
npm install
npx vercel --prod
```

---

## 💻 Rodar localmente

```bash
npm install

# Crie o arquivo .env.local:
echo "VITE_SUPABASE_URL=https://xxxx.supabase.co" > .env.local
echo "VITE_SUPABASE_ANON_KEY=eyJ..." >> .env.local

npm run dev
```

---

## 📱 Funcionalidades

| Módulo | Descrição |
|--------|-----------|
| **Dashboard** | Resumo geral: receita estimada, equipamentos alocados/disponíveis, por proprietário |
| **Equipamentos** | Cadastro completo: nome, categoria, modelo, proprietário (Hygor/ArcD), tipo de locação (diário/mensal), valor de aquisição |
| **Obras** | Cadastro de obras com status (ativa, pausada, concluída) |
| **Alocações** | Alocar, desalocar e transferir equipamentos entre obras; histórico de movimentações |
| **Relatórios** | Relatório mensal: cobranças por obra + pagamentos por proprietário; exportação Excel e PDF |
| **Config** | Cadastrar proprietários, dados da empresa, backup/restore |

---

## 🏗️ Categorias de Equipamentos

- Betoneira
- Andaime
- Materlete / Elevador
- Compactador
- Gerador
- Escora
- Esmerilhadeira
- Serra Circular
- Perfuratriz
- Compressor
- Vibrador de Concreto
- Outro

---

## 🗄️ Estrutura de dados

```json
{
  "owners": [{ "id": "hygor", "name": "Hygor", "color": "#06b6d4" }],
  "obras": [{ "id": "...", "name": "Obra 1", "status": "active" }],
  "equipamentos": [{
    "id": "...",
    "nome": "Betoneira 400L",
    "categoria": "Betoneira",
    "ownerId": "hygor",
    "tipoLocacao": "mensal",
    "valorLocacao": 800,
    "valorAquisicao": 3500
  }],
  "alocacoes": {
    "equipId": { "obraId": "obraId", "startDate": "2025-01-01" }
  },
  "movimentacoes": []
}
```

---

## 🔗 Sistema Ponto PRO

Este app é complementar ao **ArcD Ponto PRO** (gestão de funcionários). Ambos podem usar o mesmo banco Supabase, com chaves diferentes (`arced_ponto_v1` e `arced_locacoes_v1`).
