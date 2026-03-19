# ElysianOS — Sistema de Inventário

Sistema desktop para coleta e exportação de inventário de hardware e software de máquinas Windows. Gera relatórios em **HTML**, **PDF** e **XLSX**.

---

## Funcionalidades

- Coleta automática de informações do sistema:
  - Hardware (CPU, RAM, disco, GPU)
  - Software instalado
  - Rede (IPs, adaptadores)
  - Impressoras
  - Licenças Windows
  - Mapeamentos de unidades
- Relatório HTML interativo com servidor local embutido
- Exportação para Excel (.xlsx) e PDF
- Consolidador para unir múltiplos relatórios em um único XLSX
- Interface gráfica com tela de carregamento animada
- Geração de executável standalone via PyInstaller

---

## Requisitos

- Python 3.10+
- Windows

Instale as dependências:

```bash
pip install -r requirements.txt
```

---

## Como usar

```bash
python main.py
```

1. Informe o nome do colaborador na tela inicial
2. Aguarde a coleta automática dos dados
3. O relatório HTML será aberto no navegador
4. Use os botões do relatório para exportar para XLS, PDF ou abrir o consolidador

---

## Gerar executável

Execute o script `build.bat` para gerar o `.exe` com PyInstaller:

```bash
build.bat
```

O executável será gerado na pasta `dist/`.

---

## Estrutura do Projeto

```
inventario/
├── main.py                  # Ponto de entrada
├── loading_screen.py        # Tela de carregamento
├── requirements.txt
├── build.bat                # Script de build
├── assets/                  # Logo e ícones
├── collectors/              # Módulos de coleta de dados
│   ├── hardware_info.py
│   ├── software_info.py
│   ├── network_info.py
│   ├── printer_info.py
│   ├── license_info.py
│   ├── mapping_info.py
│   └── system_info.py
├── exporters/               # Módulos de exportação
│   ├── html_exporter.py
│   ├── xls_exporter.py
│   └── pdf_exporter.py
├── consolidator/            # Consolidador de relatórios
│   └── consolidator.py
└── templates/               # Templates HTML
```

---

Desenvolvido por **Luis Mesquitela**
