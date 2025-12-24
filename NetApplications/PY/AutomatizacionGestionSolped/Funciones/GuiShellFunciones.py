 # -*- coding: utf-8 -*-
import win32com.client
import sys
from typing import List, Optional

# NOTA: Para que este script funcione, se necesita una conexión activa a SAP GUI.
# El siguiente bloque es un ejemplo de cómo obtener la sesión.
# En un proyecto real, la gestión de la sesión estaría en un módulo separado.

class SapTextEditor:
      """
      Clase de abstracción para manejar el control GuiShell de tipo editor de texto en SAP.
  
      Este control no permite acceso directo a su contenido completo (.Text no existe)
      y tiene limitaciones en cómo se puede escribir en él. Esta clase implementa
      métodos robustos para leer, modificar y reescribir su contenido.
      """
      def __init__(self, session: win32com.client.CDispatch, editor_id: str):
          """
          Inicializa el wrapper del editor de texto.
  
          Args:
              session: El objeto de sesión de SAP GUI Scripting.
              editor_id: El ID completo del control GuiShell.
                         Ej: "usr/subSUBSCREEN_TEXT:SAPLSTXX:2100/cntlTEXT_EDITOR_2100/shellcont/shell"
          """
          self.session = session
          self.editor_id = editor_id
          self.editor = self._find_editor()
  
      def _find_editor(self) -> Optional[win32com.client.CDispatch]:
          """Encuentra y devuelve el objeto del editor, o None si no existe."""
          try:
              return self.session.findById(self.editor_id)
          except Exception:
              print(f"Error: No se pudo encontrar el editor con ID '{self.editor_id}'")
              return None
  
      def read_all_lines(self) -> List[str]:
          """
          Lee todo el texto del editor, línea por línea, de forma segura.
  
          La única forma fiable de saber el número de líneas es intentar leerlas
          hasta que la API COM lance una excepción, lo que indica que no hay más líneas.
  
          Returns:
              Una lista de strings, donde cada string es una línea del editor.
          """
          if not self.editor:
              return []
  
          lines = []
          line_index = 0
          while True:
              try:
                  # GetLineText es 0-indexed
                  line_text = self.editor.GetLineText(line_index)
                  lines.append(line_text)
                  line_index += 1
              except Exception:
                  # Se ha alcanzado el final del texto, la excepción es esperada.
                  break
          return lines
  
      def read_all_text(self) -> str:
          """
          Devuelve todo el contenido del editor como un único string con saltos de línea.
  
          Returns:
              El texto completo del editor.
          """
          lines = self.read_all_lines()
          return "\n".join(lines)
  
      def rewrite_text_from_lines(self, lines: List[str]):
          """
          Borra el contenido actual y escribe una nueva lista de líneas en el editor.
  
          Esta es la operación de escritura fundamental. Se realiza de forma incremental
          para evitar los problemas de `SetUnprotectedTextPart` con texto multilínea.
  
          Args:
              lines: La lista de strings que se escribirá en el editor.
          """
          if not self.editor:
              print("Error: El editor no está disponible para escribir.")
              return
  
          # 1. Limpiar el editor por completo.
          # Seleccionamos todo el texto (-1 a menudo significa "hasta el final").
          self.editor.SetSelectionIndexes(0, -1)
          self.editor.SetUnprotectedTextPart("")
  
          if not lines:
              # Si la lista está vacía, ya hemos terminado.
              return
  
          # 2. Escribir la primera línea. Esto establece el contenido inicial.
          self.editor.SetUnprotectedTextPart(lines[0])
  
          # 3. Añadir las líneas restantes una por una.
          # Este enfoque simula un "append" que el control no tiene de forma nativa.
          for i in range(1, len(lines)):
              # Obtenemos el texto actual para saber dónde posicionar el cursor.
              current_content = self.editor.GetUnprotectedTextPart()
              cursor_pos = len(current_content)
  
              # Movemos el cursor al final del todo.
              self.editor.SetSelectionIndexes(cursor_pos, cursor_pos)
  
              # Insertamos un salto de línea y la nueva línea.
              # Como la selección es un cursor (inicio=fin), SetUnprotectedTextPart actúa como una inserción.
              self.editor.SetUnprotectedTextPart("\n" + lines[i])
  
      def replace_line(self, old_line_exact: str, new_line: str, occurrences: int = -1):
          """
          Reemplaza todas las ocurrencias de una línea exacta por una nueva línea.
  
          Args:
              old_line_exact: El texto exacto de la línea a buscar.
              new_line: El nuevo texto con el que se reemplazará la línea.
              occurrences: Cuántas veces reemplazar. -1 para todas (default).
          """
          lines = self.read_all_lines()
  
          count = 0
          new_lines = []
          for line in lines:
              if line == old_line_exact and (occurrences == -1 or count < occurrences):
                  new_lines.append(new_line)
                  count += 1
              else:
                  new_lines.append(line)
  
          if count > 0:
              print(f"Se reemplazaron {count} ocurrencias de la línea.")
              self.rewrite_text_from_lines(new_lines)
          else:
              print("No se encontraron líneas que coincidieran para reemplazar.")