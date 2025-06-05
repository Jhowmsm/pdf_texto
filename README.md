# pdf_texto
extracción texto

Cómo llamar las funciones desde el config_XXXXX

XXXXX = Nombre del documento PDF

FUNCIONES
primer número despues de
    "xxxxxx:": {
      "cells": ["B11"],
      "mode": "first_number_after"
    },

Dos números después
     "PATRIMONIO": {
      "cells": ["B58", "B65"],
      "mode": "two_numbers_after"
    },

Hasta el punto

    "Auditoría:": {
      "cells": ["D19"],
      "mode": "until_dot"
    },

Hasta un salto de línea

    "Balance al 31 de diciembre de": {
      "cells": ["B56"],
      "mode": "until_newline"
    },

Texto entre dos frases o palabras claves 

     "Información clave entre frases": {
      "cells": ["B55"],
      "mode": "between_phrases",
      "start_phrase": "Opinión",
      "end_phrase": "Fundamento de la opinión"
    }



Hay varias formas de analizar los textos  
hasta un punto.
parrafo entre dos palabras
hasta un salto de linea
hasta el primer número que encuentre

y dos numéros después de


Puede localizar NIE en el documento automáticamente
