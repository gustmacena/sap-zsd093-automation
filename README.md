# SAP ZSD093 Automation

Este repositório contém um script em Python para automatizar a extração e proteção da carteira **ZSD093** no SAP. O script conecta-se ao SAP via Scripting, preenche os filtros de data, exporta os dados para um arquivo Excel e fecha o Excel automaticamente. Além disso, ele suporta extrações em lote para múltiplos períodos, exibindo um alerta ao final do processo.

## Descrição

O script realiza as seguintes tarefas:
- Conecta-se ao SAP utilizando a interface de Scripting.
- Preenche os campos de filtro (incluindo data de início e data fim) para extrair dados da carteira **ZSD093**.
- Exporta os dados extraídos para um arquivo Excel, nomeado no formato `ZSD093 - {data_inicio} - {data_fim}.xlsx`.
- Fecha automaticamente o Excel que for aberto durante o processo.
- Permite a execução de extrações com datas pré-definidas em um loop e notifica quando todas as extrações forem concluídas.

## Recursos

- **Automação do SAP:** Integração com SAP GUI Scripting.
- **Extração de dados automatizada:** Preenchimento automático dos filtros e execução do relatório.
- **Exportação de Excel:** Salva o arquivo com nomenclatura dinâmica conforme as datas informadas.
- **Fechamento automático do Excel:** Garante que o Excel seja fechado após a exportação.
- **Extração em lote:** Suporte a múltiplos períodos definidos previamente.
- **Alerta de conclusão:** Exibe uma mensagem informando que todas as extrações foram concluídas.

## Pré-requisitos

- **Sistema Operacional:** Windows.
- **SAP GUI:** Instalada com a funcionalidade de Scripting habilitada.
- **Python 3:** Instalado no sistema.
- **Pacotes Python:**
  - `pywin32` (para interação com SAP e Excel via COM)
  - `tkinter` (geralmente incluso com o Python no Windows)

Para instalar o `pywin32`, utilize:
```bash
pip install pywin32
```

## Como Utilizar

### Extração com Datas Pré-definidas (Batch)

1. Edite a lista `periodos` no script para incluir os períodos desejados, por exemplo:
   ```python
   periodos = [
       ("24.02.2025", "01.03.2025"),
       ("02.03.2025", "08.03.2025"),
       ("09.03.2025", "15.03.2025")
   ]
   ```
2. Execute o script:
   ```bash
   python nome_do_script.py
   ```
3. O script realizará a extração para cada período e, ao final, exibirá uma mensagem de alerta indicando que todas as extrações foram concluídas.

### Extração com Entrada Manual de Datas

Caso prefira solicitar as datas via interface gráfica:
1. Utilize a versão do script que solicita as datas por meio de caixas de diálogo (utilizando `tkinter`).
2. Ao executar, insira as datas quando solicitado. O script fará a extração para o período informado e perguntará se deseja realizar outra operação.

## Configuração Adicional

- **SAP GUI Scripting:**  
  Certifique-se de que o SAP GUI Scripting esteja habilitado tanto no cliente SAP quanto no servidor.
  
- **Usuário SAP:**  
  O script utiliza o usuário `ANTONIOG` por padrão. Se necessário, atualize este valor conforme sua necessidade.

- **Caminho de Salvamento:**  
  Altere o caminho onde o arquivo Excel será salvo, se necessário:
  ```python
  session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\gustavocm\\Desktop\\TesteE-commerce"
  ```

## Solução de Problemas

- **Erro "The control could not be found by id":**  
  Verifique se os identificadores dos controles no SAP (por exemplo, `"wnd[0]/usr/ctxtS_ERDAT-LOW"`) estão corretos e se a interface do SAP não foi modificada. Pode ser necessário ajustar os `time.sleep()` para dar mais tempo ao carregamento dos controles.

- **Excel não fecha:**  
  Assegure que o pacote `pywin32` esteja instalado corretamente e que não existam instâncias bloqueadas do Excel.

## Contribuição

Contribuições são bem-vindas! Se você tiver sugestões ou melhorias, por favor, abra uma issue ou envie um pull request.

## Licença

Este projeto está licenciado sob a [MIT License](LICENSE).

## Contato

*Gustavo Macena* - [gustavoaraujomacena@gmail.com](mailto:gustavoaraujomacena@gmail.com)

Projeto: [https://github.com/gustmacena/sap-zsd093-automation](https://github.com/gustmacena/sap-zsd093-automation)
```

Esse README resume as funcionalidades, requisitos e orientações para utilização e configuração do projeto. Basta ajustar as informações de contato e qualquer detalhe específico do seu ambiente.
