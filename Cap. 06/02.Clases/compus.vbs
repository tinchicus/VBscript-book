dim comput
set comput = new Compu
set sistema = CreateObject("Scripting.Dictionary")
sistema.add "so","Windows"
comput.CompuTipo = "CPU"
set comput.SistemaOperativo = sistema
wscript.echo ("CPU: " & comput.CompuTipo & vbCrLF & "SO: " & comput.getOs())
class Compu
        private tipo
        private so

        public property let CompuTipo(t)
                tipo = t
        end property

        public property get CompuTipo()
                CompuTipo = tipo
        end property

        public property set SistemaOperativo(objeto)
                set so = objeto
        end property

        public property get SistemaOperativo()
                set SistemaOperativo = so
        end property

        function getOs()
                getOs = so("so")
        end function
end class
