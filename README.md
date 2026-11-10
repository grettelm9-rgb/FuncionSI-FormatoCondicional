
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Guía de Excel | Función SI y formato condicional</title>
<style>
:root{--navy:#183153;--blue:#2563eb;--sky:#eaf2ff;--ink:#243247;--muted:#607086;--line:#dce5ef;--green:#dcfce7;--yellow:#fef3c7;--red:#fee2e2;--paper:#fff;--bg:#f3f6fb}
*{box-sizing:border-box}html{scroll-behavior:smooth}body{margin:0;background:var(--bg);color:var(--ink);font:16px/1.65 system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif}
header{background:linear-gradient(135deg,#183153,#2563eb);color:#fff;padding:42px 20px 34px}
.hero{max-width:1000px;margin:auto}.eyebrow{text-transform:uppercase;letter-spacing:.13em;font-size:.78rem;font-weight:800;opacity:.85}.hero h1{font-size:clamp(2rem,5vw,3.2rem);line-height:1.1;margin:10px 0}.hero p{max-width:700px;margin:12px 0 0;color:#e8f0ff}
nav{position:sticky;top:0;z-index:5;background:#fff;border-bottom:1px solid var(--line);box-shadow:0 3px 12px #18315310}
.navinner{max-width:1000px;margin:auto;padding:10px 18px;display:flex;gap:8px;overflow-x:auto}
nav a{white-space:nowrap;text-decoration:none;color:var(--navy);font-weight:700;font-size:.9rem;padding:7px 11px;border-radius:9px}nav a:hover{background:var(--sky)}
main{max-width:1000px;margin:28px auto;padding:0 18px 60px}.card{background:var(--paper);border:1px solid var(--line);border-radius:18px;padding:clamp(18px,3vw,28px);margin:18px 0;box-shadow:0 8px 24px #18315308}
h2{color:var(--navy);font-size:1.45rem;margin:0 0 14px}h3{color:var(--navy);margin:22px 0 8px;font-size:1.08rem}.lead{font-size:1.08rem}.callout{background:var(--sky);border-left:4px solid var(--blue);padding:14px 16px;border-radius:10px;margin:14px 0}
.formula{display:block;background:#172b4d;color:#fff;padding:14px 16px;border-radius:10px;font:600 1rem ui-monospace,SFMono-Regular,Consolas,monospace;overflow-x:auto;margin:12px 0}
table{width:100%;border-collapse:collapse;margin:14px 0;font-size:.95rem}th,td{border:1px solid var(--line);padding:10px 12px;text-align:left}th{background:#edf3fb;color:var(--navy)}tbody tr:nth-child(even){background:#fafcff}
.steps{padding-left:22px}.steps li{margin:7px 0}.tag{display:inline-block;border-radius:99px;padding:3px 10px;font-size:.82rem;font-weight:800}.green{background:var(--green);color:#166534}.yellow{background:var(--yellow);color:#92400e}.red{background:var(--red);color:#991b1b}
.practice-table input{width:100%;max-width:130px;border:1px solid #cbd5e1;border-radius:7px;padding:8px;font:inherit}.practice-table td:last-child{font-weight:700}.toolbar{display:flex;gap:10px;flex-wrap:wrap;margin:14px 0}
button{border:0;border-radius:10px;background:var(--blue);color:#fff;padding:10px 15px;font:700 .95rem system-ui;cursor:pointer}button.secondary{background:#e8eef7;color:var(--navy)}button:hover{filter:brightness(.96)}
.feedback{padding:12px 14px;border-radius:10px;margin-top:12px;display:none}.feedback.show{display:block}.good{background:var(--green);color:#166534}.bad{background:var(--yellow);color:#854d0e}
.question{padding:15px 0;border-bottom:1px solid var(--line)}.question:last-child{border:0}.question label{display:block;margin:7px 0}.question input{margin-right:8px}
footer{text-align:center;color:var(--muted);padding:20px}
@media print{nav,.toolbar,button{display:none!important}body{background:#fff}.card{box-shadow:none;break-inside:avoid}header{background:#fff;color:#183153;padding:15px 0}.hero p{color:#243247}}
</style>
</head>
<body>
<header><div class="hero">
<div class="eyebrow">Informática · Sexto grado</div>
<h1>Aprendamos Excel</h1>
<p>Guía paso a paso para comprender y practicar la <strong>función SI</strong> y el <strong>formato condicional</strong>. Lee cada explicación y después resuelve las actividades.</p>
</div></header>
<nav><div class="navinner">
<a href="#funcion-si">Función SI</a><a href="#ejemplo">Ejemplo guiado</a><a href="#formato">Formato condicional</a><a href="#actividad">Actividad práctica</a><a href="#repaso">Comprueba lo aprendido</a>
</div></nav>
<main>
<section class="card">
<h2>Antes de comenzar</h2>
<p>En esta guía aprenderás a pedirle a Excel que tome decisiones y a utilizar colores para reconocer información importante. Sigue los pasos en orden y realiza los ejemplos en un libro de Excel.</p>
<div class="callout"><strong>Tu meta:</strong> crear una tabla de calificaciones que muestre automáticamente si cada estudiante aprobó y que utilice colores para identificar sus notas.</div>
</section>

<section class="card" id="funcion-si">
<h2>1. ¿Qué es la función SI en Excel?</h2>
<p class="lead">La función <strong>SI</strong> permite que Excel tome decisiones automáticamente según una condición que nosotros establecemos.</p>
<p>Es como hacerle una pregunta a Excel. Si la respuesta es verdadera, muestra un resultado; si es falsa, muestra otro.</p>
<h3>Un ejemplo de la vida cotidiana</h3>
<p>Imagina que en una clase se establece que los estudiantes que obtengan una nota igual o superior a 70 aprueban:</p>
<ul><li>Si la nota es 70 o más, Excel escribirá <strong>«Aprobado»</strong>.</li><li>Si la nota es menor que 70, Excel escribirá <strong>«Reprobado»</strong>.</li></ul>
</section>

<section class="card">
<h2>2. ¿Cómo se escribe la función SI?</h2>
<p>La estructura de la función es:</p>
<code class="formula">=SI(condición; valor_si_verdadero; valor_si_falso)</code>
<ul>
<li><strong>Condición:</strong> es la pregunta que Excel debe comprobar.</li>
<li><strong>Valor_si_verdadero:</strong> es lo que Excel mostrará cuando la condición se cumpla.</li>
<li><strong>Valor_si_falso:</strong> es lo que Excel mostrará cuando la condición sea falsa.</li>
</ul>
<div class="callout"><strong>Consejo:</strong> en esta guía las fórmulas utilizan punto y coma ( ; ) para separar sus partes. Según la configuración de Excel, tu computadora podría utilizar comas.</div>
</section>

<section class="card" id="ejemplo">
<h2>3. Aprendamos con un ejemplo</h2>
<p>Observa esta lista de estudiantes y sus calificaciones:</p>
<table><thead><tr><th>Estudiante</th><th>Nota</th><th>Resultado</th></tr></thead><tbody>
<tr><td>Ana</td><td>95</td><td>Aprobado</td></tr><tr><td>Luis</td><td>65</td><td>Reprobado</td></tr><tr><td>María</td><td>80</td><td>Aprobado</td></tr><tr><td>Pedro</td><td>55</td><td>Reprobado</td></tr><tr><td>Sofía</td><td>70</td><td>Aprobado</td></tr>
</tbody></table>
<p>Vamos a hacer que Excel complete automáticamente la columna Resultado.</p>
<ol class="steps">
<li>Abre Excel y crea una tabla con tres columnas: <strong>Estudiante, Nota y Resultado</strong>.</li>
<li>Escribe los nombres y las notas del ejemplo anterior.</li>
<li>Haz clic en la celda <strong>C2</strong>, donde aparecerá el primer resultado.</li>
<li>Escribe esta fórmula:</li>
</ol>
<code class="formula">=SI(B2&gt;=70;"Aprobado";"Reprobado")</code>
<ol class="steps" start="5"><li>Presiona la tecla <strong>ENTER</strong>. Excel analizará la nota de Ana y escribirá «Aprobado».</li><li>Selecciona nuevamente la celda C2. Ubica el pequeño cuadrado de la esquina inferior derecha, coloca el puntero sobre él y arrastra hacia abajo hasta la última fila.</li></ol>
<p>Excel cambiará automáticamente las referencias de las celdas y calculará los resultados de todos.</p>
<h3>Otros ejemplos de la función SI</h3>
<p><strong>¿Ganó un premio?</strong> Si obtiene 90 puntos o más:</p><code class="formula">=SI(B2&gt;=90;"Ganó premio";"Sin premio")</code>
<p><strong>¿Hay existencias de un producto?</strong> Si hay más de cero unidades:</p><code class="formula">=SI(B2&gt;0;"Disponible";"Agotado")</code>
<p><strong>¿Es mayor de edad?</strong> Si tiene 18 años o más:</p><code class="formula">=SI(B2&gt;=18;"Mayor de edad";"Menor de edad")</code>
</section>

<section class="card" id="formato">
<h2>4. ¿Qué es el formato condicional?</h2>
<p>El <strong>formato condicional</strong> es una herramienta de Excel que permite cambiar automáticamente el color de las celdas según una condición.</p>
<p>Por ejemplo, podemos configurar Excel para que las notas mayores o iguales a 90 aparezcan en verde, las notas entre 70 y 89 en amarillo y las notas menores de 70 en rojo.</p>
<p>Así podemos identificar rápidamente los resultados de una tabla utilizando colores.</p>
<h3>5. ¿Cómo aplicar el formato condicional?</h3>
<ol class="steps">
<li>Selecciona todas las notas que deseas colorear. Por ejemplo, desde <strong>B2 hasta B6</strong>.</li>
<li>En la parte superior de Excel, busca la pestaña <strong>Inicio</strong>.</li>
<li>Busca la opción <strong>Formato condicional</strong>.</li>
<li>Selecciona <strong>Reglas para resaltar celdas</strong>.</li>
<li>Elige la regla que necesitas, como «Es mayor que…», «Entre…» o «Es menor que…».</li>
<li>Escribe el valor correspondiente, selecciona un color y presiona <strong>Aceptar</strong>.</li>
</ol>
<h3>Practiquemos con tres colores</h3>
<table><thead><tr><th>Condición</th><th>Color</th></tr></thead><tbody>
<tr><td>Nota mayor o igual a 90</td><td><span class="tag green">Verde</span></td></tr>
<tr><td>Nota entre 70 y 89</td><td><span class="tag yellow">Amarillo</span></td></tr>
<tr><td>Nota menor que 70</td><td><span class="tag red">Rojo</span></td></tr>
</tbody></table>
<p>Para la regla amarilla utiliza «Entre…» y escribe 70 y 89. Para la regla roja utiliza «Menor que…» y escribe 70. Para la regla verde puedes utilizar «Mayor o igual que…» y escribir 90.</p>
</section>

<section class="card" id="actividad">
<h2>6. Actividad práctica: ¡Ahora te toca a ti!</h2>
<p>Vas a crear tu propia tabla de calificaciones utilizando la función SI y el formato condicional.</p>
<h3>Parte A. Crear la tabla</h3>
<p>Abre un libro nuevo de Excel y escribe los siguientes datos:</p>
<table><thead><tr><th>Estudiante</th><th>Nota</th><th>Resultado</th></tr></thead><tbody>
<tr><td>Carlos</td><td>85</td><td></td></tr><tr><td>Daniela</td><td>95</td><td></td></tr><tr><td>Andrés</td><td>60</td><td></td></tr><tr><td>Valeria</td><td>75</td><td></td></tr><tr><td>Sebastián</td><td>50</td><td></td></tr><tr><td>Camila</td><td>100</td><td></td></tr><tr><td>Diego</td><td>68</td><td></td></tr><tr><td>Mariana</td><td>90</td><td></td></tr>
</tbody></table>
<h3>Parte B. Aplicar la función SI</h3>
<ol class="steps"><li>Haz clic en la primera celda vacía de la columna Resultado (C2).</li><li>Escribe la fórmula:</li></ol>
<code class="formula">=SI(B2&gt;=70;"Aprobado";"Reprobado")</code>
<ol class="steps" start="3"><li>Presiona ENTER.</li><li>Arrastra la fórmula hacia abajo para completar todos los resultados.</li></ol>
<h3>Parte C. Aplicar formato condicional</h3>
<ol class="steps"><li>Selecciona todas las notas (B2:B9).</li><li>Aplica color verde a las notas mayores o iguales a 90.</li><li>Aplica color amarillo a las notas entre 70 y 89.</li><li>Aplica color rojo a las notas menores de 70.</li></ol>
<h3>Parte D. Presentación</h3>
<ol class="steps"><li>Coloca los títulos en negrita.</li><li>Agrega bordes a la tabla.</li><li>Centra los datos.</li><li>Ajusta el ancho de las columnas.</li><li>Guarda el archivo con el nombre <strong>Practica_Funcion_SI_Formato_Condicional</strong>.</li></ol>
<div class="callout"><strong>Revisa tu trabajo:</strong> ¿la columna Resultado muestra Aprobado o Reprobado? ¿Las notas tienen el color que corresponde? ¿Guardaste el archivo con el nombre solicitado?</div>
</section>

<section class="card" id="repaso">
<h2>7. Comprueba lo aprendido</h2>
<p>Selecciona una respuesta para cada pregunta y luego presiona «Revisar respuestas».</p>
<form id="quiz">
<div class="question"><strong>1. ¿Para qué sirve la función SI?</strong>
<label><input type="radio" name="q1" value="a"> Para insertar imágenes.</label>
<label><input type="radio" name="q1" value="b"> Para tomar decisiones según una condición.</label>
<label><input type="radio" name="q1" value="c"> Para cambiar el tamaño de las columnas.</label></div>
<div class="question"><strong>2. ¿Qué significa la condición dentro de una función SI?</strong>
<label><input type="radio" name="q2" value="a"> La pregunta que Excel debe comprobar.</label>
<label><input type="radio" name="q2" value="b"> El color de la celda.</label>
<label><input type="radio" name="q2" value="c"> El nombre del archivo.</label></div>
<div class="question"><strong>3. ¿Qué sucede cuando una condición es verdadera?</strong>
<label><input type="radio" name="q3" value="a"> Excel elimina la fila.</label>
<label><input type="radio" name="q3" value="b"> Excel muestra el valor indicado para verdadero.</label>
<label><input type="radio" name="q3" value="c"> Excel cierra el libro.</label></div>
<div class="question"><strong>4. ¿Para qué sirve el formato condicional?</strong>
<label><input type="radio" name="q4" value="a"> Para cambiar automáticamente el formato de las celdas según reglas.</label>
<label><input type="radio" name="q4" value="b"> Para escribir una fórmula de suma.</label>
<label><input type="radio" name="q4" value="c"> Para guardar el archivo.</label></div>
<div class="question"><strong>5. ¿Cuál es la diferencia entre la función SI y el formato condicional?</strong>
<label><input type="radio" name="q5" value="a"> Ambas hacen exactamente lo mismo.</label>
<label><input type="radio" name="q5" value="b"> SI muestra un resultado según una condición; el formato condicional cambia la apariencia de las celdas.</label>
<label><input type="radio" name="q5" value="c"> El formato condicional escribe palabras y SI colorea celdas.</label></div>
<div class="toolbar"><button type="button" onclick="revisar()">Revisar respuestas</button><button class="secondary" type="reset" onclick="limpiar()">Volver a intentar</button></div>
<div id="feedback" class="feedback" role="status" aria-live="polite"></div>
</form>
</section>

<section class="card">
<h2>8. ¡Recuerda!</h2>
<ul><li>La función <strong>SI</strong> permite que Excel tome decisiones y escriba resultados automáticamente.</li><li>El <strong>formato condicional</strong> permite cambiar el aspecto de las celdas según las condiciones establecidas.</li><li>Ambas herramientas ayudan a organizar, analizar e interpretar información.</li></ul>
<p class="lead"><strong>¡Felicidades! Ahora sabes utilizar dos herramientas muy importantes de Excel.</strong></p>
<div class="toolbar"><button onclick="window.print()">Imprimir o guardar esta guía en PDF</button></div>
</section>
</main>
<footer>Guía de aprendizaje de Informática · Sexto grado</footer>
<script>
function revisar(){
 const respuestas={q1:"b",q2:"a",q3:"b",q4:"a",q5:"b"};
 let puntos=0,respondidas=0;
 for(const [pregunta,correcta] of Object.entries(respuestas)){
  const seleccion=document.querySelector('input[name="'+pregunta+'"]:checked');
  if(seleccion){respondidas++;if(seleccion.value===correcta)puntos++;}
 }
 const caja=document.getElementById("feedback");
 caja.className="feedback show "+(puntos===5?"good":"bad");
 if(respondidas<5){caja.textContent="Has respondido "+respondidas+" de 5 preguntas. Completa las que faltan. Resultado actual: "+puntos+" de 5.";}
 else if(puntos===5){caja.textContent="¡Excelente! Respondiste correctamente las 5 preguntas. ¡Muy buen trabajo!";}
 else{caja.textContent="Obtuviste "+puntos+" de 5 respuestas correctas. Vuelve a leer la guía y prueba otra vez.";}
}
function limpiar(){const caja=document.getElementById("feedback");caja.className="feedback";caja.textContent="";}
</script>
</body>
</html>
