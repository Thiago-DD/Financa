# Installer para GitHub

## Arquivo pronto

O instalador gerado localmente está em:

- `dist/FinanceiroPessoal-Setup-1.0.0.exe`

## Como subir para download no GitHub (manual)

1. Acesse o repositório no GitHub.
2. Vá em **Releases**.
3. Clique em **Draft a new release**.
4. Defina uma tag (exemplo: `v1.0.0`).
5. Arraste o arquivo `dist/FinanceiroPessoal-Setup-1.0.0.exe`.
6. Publique a release.

## Automático com GitHub Actions

Este projeto já tem workflow em:

- `.github/workflows/windows-installer.yml`

Fluxo:

1. Ao criar e enviar uma tag `v*`, o GitHub Actions compila o instalador.
2. Publica a Release com o `.exe` anexado.
