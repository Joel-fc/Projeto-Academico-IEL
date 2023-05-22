# Sistema de Solicitação de Ordens de Serviço aos Ferramenteiros.
  Este é um projeto acadêmico desenvolvido como parte de uma Jornada de Aprendizagem promovida pela instituição "Faculdades da Indústria (IEL)", em parceria com a Volkswagen do Brasil. Seu objetivo principal é solucionar o problema de falta de padronização e rastreamento das solicitações de ordens de serviço direcionadas aos ferramenteiros.

  Tradicionalmente, essas solicitações eram feitas de forma verbal ou através de documentos físicos, chats e outros meios não padronizados. Portanto, buscamos apresentar uma solução que pudesse ser implementada com o menor custo possível, considerando o desenvolvimento, implementação e futuras manutenções.

  Com base nessas premissas, optamos por desenvolver o software utilizando a linguagem Visual Basic for Applications (VBA) e o banco de dados ACCESS para o cadastro de usuários e armazenamento das ordens de serviço. Essa escolha foi embasada no fato de que ambas as ferramentas fazem parte do pacote Office, amplamente utilizado no mercado de trabalho.

# Aviso Legal
**Importante**: Este projeto e os problemas abordados nele não são de forma alguma relacionados à empresa Volkswagen. Este projeto foi desenvolvido como um trabalho acadêmico para a disciplina de Engenharia de Software da instituição Faculdades da Indústria - IEL. Os problemas e necessidades mencionados são fictícios e foram criados apenas para fins educacionais e de aprendizado.

O objetivo deste projeto foi desenvolver um sistema de rastreamento e organização de ordens de serviço, baseado em VBA e utilizando o banco de dados ACCESS, como uma solução hipotética para melhorar a gestão das solicitações de ordens de serviço.

Por favor, tenha em mente que este projeto não tem nenhuma relação com a empresa Volkswagen e não deve ser interpretado como uma representação precisa ou uma solução oficial adotada pela empresa. É apenas um trabalho acadêmico e todas as informações e dados utilizados são fictícios.

# Funcionalidades
  * Tela de Login: Permite que os usuários façam login no sistema, com opções para cadastrar uma nova conta ou recuperar a senha, caso necessário. As senhas recuperadas são enviadas por e-mail.
  * Emissão de Ordens de Serviço: Os usuários têm acesso a uma tela com campos para preenchimento das informações necessárias para a emissão de uma ordem de serviço aos ferramenteiros. Após preencher os campos e enviar a O.S através do botão "Enviar Ordem", o gestor responsável recebe um e-mail com um PDF da ordem para aprovação ou reprovação.
  * Aprovação de Ordens de Serviço: Os usuários têm acesso a uma tela específica onde podem visualizar todas as ordens de serviço inseridas e verificar o status de cada uma delas. O gestor, por sua vez, pode analisar cada ordem e tomar a decisão de aprová-la ou reprová-la. Essa ação acarreta no envio de um e-mail contendo o PDF correspondente ao ferramenteiro responsável.
  * Relatório de Ordens de Serviço: Os usuários têm acesso a uma tela para gerar relatórios de todas as ordens de serviço inseridas. É possível aplicar filtros por data, usuário que inseriu a ordem e status da ordem (aprovada, reprovada ou em análise).

# Pré-requisitos
  * Sistema operacional Windows 7 ou superior.
  * Microsoft Excel 2016 ou versão superior.
  * Microsoft Access 2016 ou versão superior.
  * Habilitação dos seguintes suplementos no Visual Basic for Applications (VBA):
    * Visual Basic for Applications
    * Microsoft Excel 16.0 Object Library
    * OLE Automation
    * Microsoft Office 16.0 Object Library
    * Microsoft Forms 2.0 Object Library
    * Microsoft ActiveX Data Objects 6.1 Library

# Configurações Adicionais
  Para habilitar o envio de e-mails, é necessário realizar algumas configurações adicionais no código VBA. No módulo "modEnviaEmail", dentro das Subs "emailOrdem" e "emailAprovacaoOrdem", é necessário fornecer o endereço de e-mail para o qual as ordens de serviço devem ser enviadas.

  Além disso, no formulário "FormAprovacaoOrdem", na Sub "btnSalvaAlteracaoStatus_Click()", é necessário atribuir um endereço de e-mail à variável "usuarioComPermissao". Esse endereço de e-mail será responsável pela alteração do status das ordens.

  Certifique-se de realizar essas configurações corretamente para garantir o funcionamento adequado do sistema de envio de e-mails.

# Como utilizar
  * Quando você baixar o arquivo de Excel, por motivos de segurança, a Microsoft desabilita automaticamente as macros criadas por terceiros. Portanto, você precisará fazer a liberação manualmente seguindo as etapas abaixo:
    * Clique com o botão direito do mouse no arquivo de Excel e selecione a opção "Propriedades".
    * Na janela que abrir, localize a caixa de seleção no canto inferior direito com a descrição "Desbloquear".
    * Marque essa caixa de seleção e clique em "Aplicar" e, em seguida, em "OK".
    * Feito isso, você poderá abrir o arquivo normalmente no Excel.

    Esse processo de liberação de macros só precisa ser realizado uma única vez. Após desbloquear o arquivo, as macros estarão habilitadas e você poderá utilizá-las sem problemas.

  * Certifique-se de ter o Microsoft Excel e o Microsoft Access instalados e licenciados em sua máquina.
  * Habilite os suplementos necessários no Visual Basic for Applications (VBA).
  * Coloque o arquivo de Excel e o banco de dados no mesmo diretório. Caso opte por utilizar diretórios diferentes, será necessário modificar o código de conexão com o banco.
  * Configure as opções de envio de e-mails no código VBA, conforme as instruções fornecidas.
  * Execute o projeto através do arquivo de Excel e utilize as funcionalidades disponíveis.
  * Usuário:
    * Login: ADMIN@ADMIN.COM
    * Senha: ADMIN

# Limitações
  O projeto não foi desenvolvido com responsividade em mente, podendo gerar conflitos em diferentes resoluções de tela.

# Observações
  Para fins de testes, o banco de dados disponibilizado possuí algumas ordens de serviços fictícias, facilitando os testes de relatórios e aprovação/reprovação das O.S.
  
## Reconhecimentos

Agradeço a todos os membros da minha equipe que contribuíram para o desenvolvimento e aprimoramento deste projeto. Suas idéias, esforços e trabalho em equipe foram fundamentais para o sucesso do projeto.

Equipe de Desenvolvimento:
  - Joel França da Cruz (https://github.com/Joel-18-And-Life)
  - Murilo Vieira Santos Maria (https://github.com/MiloVSM)

Equipe de Documentações:
  - Arthur Ruiz Garcia (https://github.com/ydrozzy0)
  - Jairo Marcos do Nascimento Santos Filho (https://github.com/Jaironascimentof)
  - Ana Carolina Fernandes (https://www.linkedin.com/in/ana-carolina-fernandes-3b37981a9/)

Também gostaria de agradecer aos professores e orientadores que nos apoiaram ao longo do processo e nos forneceram orientações valiosas:
  - Cassiana Fagundes da Silva (https://www.linkedin.com/in/cassianafagundesdasilva/)
  - Fabio Garcez Bettio (https://www.linkedin.com/in/fabio-bettio/)

## Contato

  - E-mail: joel.pessoal.cruz@gmail.com
  - LinkedIn: https://www.linkedin.com/in/joel-cruz-a814b71a2/


Fique à vontade para entrar em contato comigo caso tenha alguma pergunta, sugestão ou feedback relacionado a este projeto. Estou disponível para ajudar e discutir possíveis melhorias.








