#Origen
El script se desarrolló con el fin de obtener una calificación a estudiantes
que se les ha recomendado un conjunto de habilidades en la plataforma 
[Khan Academy](https://es.khanacademy.org/).

#Requisitos
Entorno donde pueda correr python, con la librería _openpyxl_
#Modo de uso
Basicamente se debe correr el script en el directorio donde se halla el archivo 
de configuración _"datos.xlsx"_ y el documento que se descarga con la información
de los estudiantes desde la plataforma [Khan Academy](https://es.khanacademy.org/), 
el nombre de tal archivo se debe especificar en la hoja denominada datos.

		$python califica.py 

#Configuración
La hoja de cálculo _"datos.xlsx"_ es donde se halla la configuración de como 
generar los reportes, consta de cuatro hojas:

- datos: están los nombres de los reportes de khan vinculados con grupos o clases,
 además se indica que periodo o grupo de habilidades calificar.
- grupos: cada columna contiene los códigos de estudiantes (únicos) de un grupo.
El nombre del grupo debe terminar en un número, pues así el script halla el número
de grupo correcto según la especificación en la parte de _datos_.
- habilidades: Listado de habilidades en diferentes grupos, pues en la práctica
del autor, se recomiendan grupos de habilidades en diferentes periodos de tiempo a
los estudiantes.
- calificación:Es donde se asigna la convención de la nota numércia en base a 
la nota cualitativa que da la plataforma Khan Academy.


#Recomendaciones
En la plataforma Khan Academy se debe tener a los estudiantes identificados por
código, en caso de que no estén así, es posible renombrarlos manualmente. Aunque
es posible usar nombres pero quizá no es muy práctico a nivel de tiempos en que 
se ejecuta el script, y posibles problemas por homónimos.


