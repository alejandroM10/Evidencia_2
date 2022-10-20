import datetime
import openpyxl
import csv

# Creación de diccionarios 
horario_L= []
disponible= []
clientes_D= {}
salas_D= {}
reservaciones_D= {}
turno_D= {1: "Matutino",
          2: "Vespertino",
          3: "Nocturno"}
while True:
    # Menu principal
    print("MENÚ PRINCIPAL")
    print("[1]. Reservaciones")
    print("[2]. Reportes")
    print("[3]. Registrar una sala")
    print("[4]. Registrar a un nuevo cliente")
    print("[5]. Salir") 
    opcion=input("¿Que desea hacer?: ")
    # Validar que selecciono una opcion del menu correcta
    if (not opcion in "12345"):
        print("Opción no valida, ingresa una opcion del menú")
        continue

    if opcion=="1": # Reservaciones
        # submenu
        print("SUBMENÚ")
        print("[1]. Registrar nueva reservación")
        print("[2]. Modificar descripción de una reservación")
        print("[3]. Consultar disponibilidad de salas para una fecha")
         
        opcion2=input("¿Que desea hacer?: ")
        if (not opcion in "123"):
            print("Opción no valida, ingresa una opcion del menú")
            continue
        if opcion2=="1": # Registrar reservación
            clave_cliente = int(input("Ingresa la clave del cliente: "))
            if (not clave_cliente in clientes_D.keys()):
                print("Cliente no registrado. Registrarse primero")
            else:
                folio = max(reservaciones_D, default=0)+1
                nombre_evento=input("Ingresa el nombre del evento: ")
                # Verificar fecha
                Fecha_actual=datetime.date.today()
                fecha_reservacion=input("Ingresa la fecha a reservar: ")
                Fecha_reserva=datetime.datetime.strptime(fecha_reservacion, "%d/%m/%Y" ).date()
                mes1=Fecha_reserva.month  
                mes2=Fecha_actual.month
                dias_aprox=Fecha_reserva.day - Fecha_actual.day
                if dias_aprox <= 2 and mes1 <= mes2:
                        print("La reservacion se tiene que hacer con 2 dias de anticipacion")
                else: 
                    if Fecha_reserva in horario_L: # cuando la fecha existe
                        print(salas_D)
                        sala=input("Selecciona la sala: ")
                        print(turno_D)
                        turno=input("Selecciona el turno: ")
                        lugarfecha= [(sala,turno)]
                        if lugarfecha in disponible:
                            print("Fecha ocupada")
                            print("\n")
                        else: 
                            disponible.append(lugarfecha)
                            print("Reservacion registrada")
                            reservaciones_D.update({folio: [folio, nombre_cliente, nombre_evento,Fecha_reserva, turno,sala]})
                            print("\n")    
                    else: # cuando la fecha no existe
                        horario_L.append(Fecha_reserva)
                        print(salas_D)
                        sala=input("Selecciona la sala: ")
                        print(turno_D)
                        turno=input("Selecciona el turno: ")
                        lugarfecha= [(sala,turno)]
                        disponible.append(lugarfecha)
                        print("Reservacion registrada")
                        reservaciones_D.update({folio: [folio, nombre_cliente, nombre_evento,Fecha_reserva, turno,sala]})
                        print("\n")
        elif opcion2=="2": # Modificar descripción de una reservación solo el nombre del evento 
            Modificar= int(input("Ingresa el folio de la reservación: "))
            cambio=reservaciones_D.get(Modificar)
            if cambio==None:
                print("Reservación no encontrada")
            else: 
                print("Datos Actuales:", {cambio[0]},{cambio[1]},{cambio[2]},{cambio[3]},{cambio[4]})
                nombre_nuevo=input("Ingresa el nuevo nombre del evento: ")
                nombre_evento=nombre_nuevo
                reservaciones_D.update({folio:[folio,nombre_cliente,nombre_evento,fecha_reservacion,turno,sala]})
            print("\n")
        else: # Consultar disponibilidad de salas para una fecha
            buscar_fecha2=input("Ingresa la fecha para ver las reservaciones: ")
            disp=datetime.datetime.strptime(buscar_fecha2,"%d/%m/%Y").date()
            if disp in horario_L:
                print("Salas existentes: Sala, Turno")
                print(f"{sala},1")
                print(f"{sala},2")
                print(f"{sala},3")
                print(f"Salas ocupadas: {lugarfecha}")

    elif opcion=="2": # Reportes
        # submenu
        print("SUBMENÚ")
        print("[1]. Reporte en pantalla para una fecha")
        print("[2]. Exportar reporte tabular en excel")
        opcion3=input("¿Que desea hacer?: ")
        if (not opcion in "12"):
            print("Opción no valida, ingresa una opcion del menú")
            continue
        buscar_fecha=input("Ingresa la fecha para ver las reservaciones: ")
        buscar_fecha=datetime.datetime.strptime(buscar_fecha,"%d/%m/%Y").date()
        if opcion3=="1": # Reporte en pantalla para una fecha
            for Fecha in reservaciones_D.keys():
                for valores in reservaciones_D.values():
                    print("*"*80)
                    print("*"*20,f"Reporte de reservaciones para el dia {buscar_fecha}","*"*20)
                    print("*"*90)
                    print("*"*90) 
                    print("{:<15} {:<15} {:<15} {:<15}".format("Sala","Cliente","Evento","Turno"))
                    print("*"*90)
                    print(f"{valores[5]}\t\t\t\t{valores[1]}\t\t\t{valores[2]}\t\t\t{valores[4]}")
                    print("*"*35,"Fin del reporte","*"*35)
                    print("\n")
                    break
        else: # Exportar reporte tabular en excel
            libro = openpyxl.Workbook()
            hoja = libro["Sheet"] 
            hoja.title = "Reporte"
            hoja["E1"].value = "Reporte de reservaciones"
            hoja["E3"].value = "Sala,Cliente,Evento,Turno"
            for Fecha in reservaciones_D.keys():
                for valores in reservaciones_D.values():
                    valor1=valores[5]
                    valor2=valores[1]
                    valor3=valores[2]
                    valor4=valores[4]
                    hoja["D5"].value= valor1
                    hoja["E5"].value= valor2
                    hoja["F5"].value= valor3
                    hoja["G5"].value= valor4
                    break
                libro.save(f"Reporte{buscar_fecha}.xlsx")
                print("Reporte creado exitosamente!")

    elif opcion=="3": # Registrar una sala
        clave_sala= max(salas_D,default=0)+1
        while True:
            nombre_sala= input("Ingresa el nombre de la sala: ")
            if (nombre_sala==""):
                print("El nombre no se puede omitir")
                continue
            else:
                while True:
                    cupo_sala= input("Ingresa el cupo de la sala: ")
                    if (cupo_sala==""):
                        print("El dato no se puede omitir")
                        continue   
                    break
                break            
        salas_D.update({clave_sala:[clave_sala,nombre_sala,cupo_sala]})
        print("Sala agregada")
        print("\n")

    elif opcion=="4": # Registrar un cliente
        clave_cliente= max(clientes_D, default=0)+1
        # Captura el nombre. Si se omite se envia un mensaje indicando que no debe omitirse
        while True: 
            nombre_cliente= input("Ingresa el nombre del cliente: ")
            if (nombre_cliente==""):
                print("El nombre no se puede omitir")
                continue
            break
        clientes_D.update({clave_cliente:[clave_cliente, nombre_cliente]})
        print("Cliente agregado")
        print("\n")

    else:
        reservaciones_CE = reservaciones_D
        #Paso 3: Abrir, en modo de escritura, el archivo destino
        with open("Reservaciones.csv","w", newline="") as archivo:
            #Paso 4: Establecer una salida de escritura
            grabador = csv.writer(archivo)
            #Paso 5: Grabar el encabezado (OPCIONAL)
            grabador.writerow(("Clave", "folio","nombre_cliente","nombre_evento","fecha_reservacion","turno","sala"))
            #Paso 6: Iterar sobre los elementos de los datos a grabar o bien pedir de golpe que se graben todos los elementos
            grabador.writerows([(clave, datos[0], datos[1], datos[2], datos[3], datos[4], datos[5]) for clave, datos in reservaciones_CE.items()])
        
        clientes_CE = clientes_D
        with open("Clientes.csv","w", newline="") as archivo2:
            #Paso 4: Establecer una salida de escritura
            grabador2 = csv.writer(archivo2)
            #Paso 5: Grabar el encabezado (OPCIONAL)
            grabador2.writerow(("Clave","Clave_cliente","nombre_cliente"))
            #Paso 6: Iterar sobre los elementos de los datos a grabar o bien pedir de golpe que se graben todos los elementos
            grabador2.writerows([(clave2, datos2[0], datos2[1]) for clave2, datos2 in clientes_CE.items()])
        
        salas_CE = salas_D
        with open("Salas.csv","w", newline="") as archivo3:
            #Paso 4: Establecer una salida de escritura
            grabador3 = csv.writer(archivo3)
            #Paso 5: Grabar el encabezado (OPCIONAL)
            grabador3.writerow(("Clave","Clave_sala","nombre_sala","cupo_sala"))
            #Paso 6: Iterar sobre los elementos de los datos a grabar o bien pedir de golpe que se graben todos los elementos
            grabador3.writerows([(clave3, datos3[0], datos3[1], datos3[2]) for clave3, datos3 in salas_CE.items()])
        
        print("Ejecución Finalizada")
        break
   