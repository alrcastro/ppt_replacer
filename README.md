# PPT Replacer
###### PT/BR
Essa é uma biblioteca simples em javascript cujo o objetivo é exportar uma apresentação usando um template como base e um array de objetos javascript como modelo.

### O Template
[Imagem](help/ppttemplate.png)
O template é um slide aonde ele basicamente substituirá o texto (que vai ser o nome das propriedades de seu objeto) pelo valor. Você também pode usar imagens no seu template, mas para o código entender qual o nome da imagem você precisa colocar um *Texto Alt* nele:
[TextoAlt](help/textalt.png)

### O Código
O objeto principal precisa seguir essas condições:
Os dados devem estar dentro de 'data' e as imagens dentro de 'imgData'. 
Por exemplo:  Se no template você quer substituir os valores de varNome e varDescricao você teria esse objeto:
...
 let obj = { data: [ 
					 {varNome: 'Jose', varDescription: 'CEO'},
					 {varNome: 'Raimundo', varDescription: 'Professor'}
					 ]};
...
A propriedade = Nome que deseja substituir e o valor = novo valor.
Depois basta instanciar a classe e chamar o método loadAndProcess passando o path do template e seu objeto:
...
	const pptTemplate = 'assets/job_info.pptx';
    let pptx = new Presentation();
	pptx.loadAndProcess(obj, pptTemplate).then(() => { pptx.downloadBuffer() });
...
Ele retorna uma promise e depois basta chamar o downloadBuffer para fazer o download do ppt

### Sample/Exemplo
Library that creates an PPT presentation based on a template and a list of objects (in javascript)
You can try at https://cleo-209421.firebaseapp.com/ 
