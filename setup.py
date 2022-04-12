from cx_Freeze import setup, Executable

setup(name = "Preenche PDF",
    version = "1.1",
    description = "O executável preenche o word e gera o pdf com o nome dos colaboradores passado na lista do excel.",
    executables = [Executable("preenche_pfd.py",icon="logoc.ico")]
         )