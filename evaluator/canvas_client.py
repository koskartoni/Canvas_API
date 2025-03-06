from config.settings import API_URL, API_KEY
from canvasapi import Canvas
import pandas as pd

# Conectar con Canvas
canvas = Canvas(API_URL, API_KEY)

def obtener_cursos():
    """Obtiene la lista de cursos del usuario autenticado."""
    try:
        cursos = canvas.get_current_user().get_courses()
        return {curso.id: curso.name for curso in cursos}
    except Exception as e:
        print(f"Error al obtener los cursos: {e}")
        return {}

def obtener_alumnos(curso_id):
    """Obtiene la lista de alumnos de un curso específico."""
    try:
        curso = canvas.get_course(curso_id)
        alumnos = curso.get_users(enrollment_type=['student'])
        return [{"Estudiante ID": alumno.id, "Nombre": alumno.name} for alumno in alumnos]
    except Exception as e:
        print(f"Error al obtener alumnos del curso {curso_id}: {e}")
        return []

def obtener_calificaciones(curso_id):
    """Obtiene las calificaciones de los estudiantes en un curso."""
    try:
        curso = canvas.get_course(curso_id)
        asignaciones = curso.get_assignments()
        datos = []

        for asignacion in asignaciones:
            for envio in asignacion.get_submissions():
                if envio.score is not None:
                    datos.append({
                        "Curso": curso.name,
                        "Asignación": asignacion.name,
                        "Estudiante ID": envio.user_id,
                        "Calificación": envio.score
                    })

        return pd.DataFrame(datos)
    except Exception as e:
        print(f"Error al obtener calificaciones del curso {curso_id}: {e}")
        return pd.DataFrame()

if __name__ == "__main__":
    cursos = obtener_cursos()
    print("Cursos disponibles:")
    for id_curso, nombre in cursos.items():
        print(f"{id_curso}: {nombre}")

    try:
        curso_id = int(input("Ingrese el ID del curso para extraer calificaciones: "))
        if curso_id not in cursos:
            raise ValueError("El ID del curso no es válido.")

        # Guardar el ID del curso seleccionado
        with open("selected_course.txt", "w") as f:
            f.write(str(curso_id))

        print(f"Curso seleccionado guardado: {curso_id}")

        # Obtener y guardar alumnos
        alumnos_data = obtener_alumnos(curso_id)
        if alumnos_data:
            df_alumnos = pd.DataFrame(alumnos_data)
            df_alumnos.to_csv("alumnos.csv", index=False)
            print("Archivo 'alumnos.csv' generado con éxito.")

        # Obtener y guardar calificaciones
        df_calificaciones = obtener_calificaciones(curso_id)
        if not df_calificaciones.empty:
            df_calificaciones.to_csv("calificaciones.csv", index=False)
            print("Archivo 'calificaciones.csv' generado con éxito.")
        else:
            print("No se encontraron calificaciones.")

    except ValueError as e:
        print(f"Error: {e}")
