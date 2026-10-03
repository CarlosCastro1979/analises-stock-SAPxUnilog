# Aulas PT

App de telemóvel (Expo / React Native) para duas pessoas marcarem **horários feitos** no calendário, com preço fixo por aula e resumo pronto para WhatsApp.

## Como usar

1. Instalam a mesma app nos dois telemóveis.
2. Na primeira abertura metem o mesmo **código da casa**.
3. Cada um marca os seus dias; o total do mês = aulas dos dois × preço (definições).
4. O botão **Enviar no WhatsApp** abre o texto já preenchido.

## Pré-requisito Supabase

Corre uma vez o SQL em [`supabase.sql`](./supabase.sql) no SQL Editor do projeto Supabase (o mesmo da Performance Logística). Sem isto a app não consegue gravar.

## Desenvolvimento

```bash
cd aulas-pt
npm install
npm start
```

Depois abre no **Expo Go** (QR code) ou num emulador.

```bash
npm run typecheck
npm run test:logic
npx expo export --platform android   # smoke bundle
```

## APK (Android)

Com conta Expo / EAS:

```bash
npx eas-cli@latest build --platform android --profile preview
```

Ou, com toolchain Android local: `npx expo run:android` / gerar release APK.

## Estrutura

- `src/app/` — rotas Expo Router
- `src/screens/` — ecrãs (código da casa, calendário, definições)
- `src/hooks/useAulasData.ts` — Supabase + realtime + autosave
- `supabase.sql` — tabelas, RLS e realtime
