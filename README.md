# Robô Relatórios Recupera

Esse script foi desenvolvido originalmente para extração de relatórios de comissionamento, mas pode futuramente ser modificado para extração de quaisquer outros relatórios no Recupera.

O script roda 99% automatizado com Selenium, mas devido a uma particularidade do Recupera **a seguinte linha precisa ser modificada MANUALMENTE com as coordenadas X, Y do mouse para clicar na tela de preenchimento do relatório**:

~~~python
...
# pyautogui assume para preencher campos que não foram possíveis preencher com selenium
pa.click(X, Y)
....
~~~

Para pleno funcionamento do script você precisa:
1. Ter o Python instalado
2. Instalar as dependências do projeto
3. Ter o Chrome instalado
4. Usar a aplicação MouseInfo para pegar a coordenada do seu mouse na página do Recupera
5. Colar a coordenada do mouse na linha 104 do script Python

# Passo a passo do zero (caso precise):
1. Realize o download do Python em https://www.python.org/downloads/ e o instale no computador.
2. Com o Python instalado, instale as dependências necessárias com o seguinte comando no CMD:
~~~cmd
pip install -U customtkinter mouseinfo pyautogui selenium webdriver_manager
~~~
3. Caso não consiga realizar o download dessa forma, vá até o local de instalação do Python, acesse a pasta Scripts:
   
![image](https://github.com/Cauamattosprj/automacao-recupera-cobrart/assets/71733712/430eba82-cdb8-4e6e-828c-bccd4d54fe04)

5. Clique na barra de endereços e escreva **CMD**
   
![image](https://github.com/Cauamattosprj/automacao-recupera-cobrart/assets/71733712/1e7f37ec-174d-4454-bd2d-46da3725b412)

7. Realize o passo 2 novamente.
8. Com todas as dependências instaladas, rode o seguinte comando no CMD:
~~~cmd
py
from mouseinfo import mouseInfo
mouseInfo()
~~~
7. Com a tela do MouseInfo aberta, abra o **Google Chrome** realize login no site do Recupera, **maximize a tela com F11** e clique para gerar planilha do Excel:
   
![image](https://github.com/Cauamattosprj/automacao-recupera-cobrart/assets/71733712/20e6188e-7e25-402e-81cc-54d30596a6dd)

8. Agora use o MouseInfo para pegar a coordenada do mouse, conforme o print abaixo:

![image](https://github.com/Cauamattosprj/automacao-recupera-cobrart/assets/71733712/280e9f4a-b8eb-4535-8cf5-1761eacb261d)

9. Cole neste trecho do código e tudo pronto.

![image](https://github.com/Cauamattosprj/automacao-recupera-cobrart/assets/71733712/97e560ee-7d88-4a64-a9e5-ce420483cc1c)
