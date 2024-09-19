# OfficePwn

Este projeto consiste em um script VBA que pode ser integrado a documentos Microsoft Office (.docx, .xlsm, .pptx). Ele executa uma enumeração de informações do sistema e abre uma **reverse shell** para o atacante, além de tentar enviar um arquivo de log para um compartilhamento de rede remoto.

⚠️ **AVISO**: Este projeto é destinado apenas a fins educacionais e para testes de segurança autorizados. O uso indevido deste código para comprometer sistemas sem consentimento é **ilegal** e **antiético**. Use-o apenas em ambientes controlados onde você tenha permissão explícita para realizar testes de segurança.

## Funcionalidades

- Enumeração de informações do sistema, incluindo:
  - Sistema operacional
  - Endereços IP
  - Serviços em execução
  - Portas abertas
- Reverse shell que permite ao atacante executar comandos remotamente
- Envio do arquivo de log de enumeração para um compartilhamento de rede controlado pelo atacante

## Requisitos

- Microsoft Office (Word, Excel ou PowerPoint) com suporte a Macros
- Windows PowerShell habilitado no alvo
- Acesso a uma máquina atacante para ouvir a conexão reverse shell

## Instalação no Alvo

### Configurando o Documento Office:

1. Abra o Microsoft Word, Excel ou PowerPoint no alvo.
2. Acesse a aba **Desenvolvedor**:
   - No Office, vá em **Arquivo** > **Opções** > **Personalizar Faixa de Opções** e ative a aba **Desenvolvedor**.
3. Clique em **Visual Basic** para abrir o editor de VBA.
4. No painel esquerdo, clique com o botão direito em **ThisWorkbook** (para Excel) ou **ThisDocument** (para Word) e selecione **Exibir Código**.
5. Cole o código VBA contido neste repositório no editor.
6. Modifique o IP e porta na função `CriarReverseShell` para corresponder ao IP e porta da máquina atacante.
7. (Opcional) Renomeie funções e macros para algo mais discreto, como `GenerateReport` ou `UpdateData`.
8. Feche o editor de VBA e salve o documento como um arquivo habilitado para macros:
   - Para Word, salve como `.docm`.
   - Para Excel, salve como `.xlsm`.
   - Para PowerPoint, salve como `.pptm`.

### Configuração do Atacante:

1. Na máquina atacante, configure um listener para receber a conexão da reverse shell:
   ```bash
   nc -lvp 4444
   ```
   - Substitua a porta `4444` pela porta especificada no script VBA.
   
2. Certifique-se de que o IP do servidor atacante esteja acessível pela máquina vítima.

## Uso

1. Envie o documento malicioso para o alvo (por exemplo, via phishing ou engenharia social).
2. O usuário abre o documento e executa a macro, o que dispara o código.
3. A macro:
   - Coleta informações do sistema e as salva em um arquivo de log (`enum_resultados.txt`) na área de trabalho da vítima.
   - Tenta enviar esse arquivo de log para um compartilhamento de rede configurado no servidor atacante.
   - Estabelece uma reverse shell, permitindo ao atacante enviar comandos remotamente para a máquina alvo.
4. O atacante, conectado via **netcat** ou outro listener, pode controlar o sistema infectado e visualizar os resultados no arquivo de log.

## Estrutura do Projeto

```bash
.
├── README.md            # Este arquivo
├── LICENSE              # Licença MIT para o projeto
└── vba_reverseshell.vba # Código VBA com Reverse Shell e Enumeração
```

## Contribuindo

Sinta-se à vontade para enviar **pull requests** ou abrir **issues** para melhorias no código ou na documentação.

## Licença

Este projeto está licenciado sob a Licença MIT - consulte o arquivo [LICENSE](./LICENSE) para mais detalhes.
