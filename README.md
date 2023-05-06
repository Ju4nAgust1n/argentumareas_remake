Rediseño del sistema de áreas de Argentum Online

Desde hace ya algunos años dicho MMORPG cuenta con un sistema que divide el mapa en áreas para hacer más eficiente el envío de información a los clientes. El asunto es que dicha implementación, que viene por defecto con el juego, puede hacerse aún más eficiente.

Este nuevo sistema posee las siguientes ventajas, a criterio de su autor:

1) Rediseño de lógica

![image](https://user-images.githubusercontent.com/34247356/236646045-c3f39914-3460-49e5-a218-ca60b45aef48.png)

![image](https://user-images.githubusercontent.com/34247356/236646115-1dd565e8-eec1-4584-b9b7-6b0ef6168c3d.png)

2) Es modificable desde archivos .ini, .dat (véase los archivos: areas.dat y areaspos.ini incluidos en la carpeta 'dats' del servidor), y constantes de código.

3) Se utiliza Microsoft Scripting Runtime para:
  *Manejar de manera más clara las colas de usuarios, npcs y objetos por área,
  *Identificarlos a cada uno de ellos con una clave específica,
  *Ahorrar líneas de código.

