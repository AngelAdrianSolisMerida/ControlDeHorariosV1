def generar_horarios(self):
    if not self.datos_procesados is None and self.validar_fechas():
        try:
            tipo_examen = self.exam_type.get()
            
            # Configuración según tipo de examen
            config = {
                "Ordinario": {
                    "horas": 3,
                    "examenes_por_turno": 2,
                    "dias_recomendados": 5,
                    "horario_matutino": ["08:00", "11:00"],
                    "horario_vespertino": ["15:00", "18:00"]
                },
                "Extraordinario": {
                    "horas": 2,
                    "examenes_por_turno": 3,
                    "dias_recomendados": 3,
                    "horario_matutino": ["08:00", "10:00", "12:00"],
                    "horario_vespertino": ["15:00", "17:00", "19:00"]
                }
            }[tipo_examen]
            
            # Generar días hábiles (Lunes a Viernes)
            dias_habiles = []
            current_date = self.fecha_inicio
            while current_date <= self.fecha_fin:
                if current_date.weekday() < 5:  # 0=Lunes, 4=Viernes
                    dias_habiles.append(current_date)
                current_date += timedelta(days=1)
            
            if not dias_habiles:
                messagebox.showerror("Error", "No hay días hábiles en el rango de fechas seleccionado")
                return
            
            # Verificar días mínimos recomendados
            if len(dias_habiles) < config["dias_recomendados"]:
                messagebox.showwarning(
                    "Advertencia",
                    f"Se recomiendan al menos {config['dias_recomendados']} días hábiles para exámenes {tipo_examen}.\n"
                    f"Actualmente hay {len(dias_habiles)} días hábiles en el rango seleccionado."
                )
            
            # Preparar datos del Excel
            df = self.datos_procesados.copy()
            
            # Limpieza y validación de datos
            df = df.dropna(subset=['SEMESTRE', 'MATERIA', 'GRUPO', 'TURNO', 'DOCENTE'])
            
            # Normalizar datos
            df['SEMESTRE'] = df['SEMESTRE'].astype(str).str.strip()
            df['GRUPO'] = df['GRUPO'].astype(str).str.strip().str.upper()
            df['TURNO'] = df['TURNO'].astype(str).str.strip().str.upper()
            df['DOCENTE'] = df['DOCENTE'].astype(str).str.split(',').str[0].str.strip()
            
            # Identificar optativas
            df['ES_OPTATIVA'] = df['MATERIA'].str.contains('OPTATIVA', case=False, na=False)
            
            # Verificar si hay datos después de limpieza
            if df.empty:
                messagebox.showerror("Error", "No hay datos válidos después de la limpieza")
                return
            
            # Obtener todos los semestres y grupos únicos
            todos_semestres = sorted(df['SEMESTRE'].unique(), key=lambda x: int(x) if x.isdigit() else x)
            todos_grupos = sorted(df['GRUPO'].unique())
            
            # Organizar materias por semestre y grupo
            materias_por_grupo = {}
            for semestre in todos_semestres:
                materias_por_grupo[semestre] = {}
                for grupo in todos_grupos:
                    materias_grupo = df[(df['SEMESTRE'] == semestre) & (df['GRUPO'] == grupo)]
                    if not materias_grupo.empty:
                        materias_por_grupo[semestre][grupo] = materias_grupo
            
            # Estructura para controlar asignaciones
            docentes_asignados = {}  # {fecha: {docente: [horarios]}}
            grupos_asignados = {}    # {fecha: {grupo: [horarios]}}
            
            # Función para verificar disponibilidad de docente
            def docente_disponible(docente, fecha, hora):
                if fecha not in docentes_asignados:
                    return True
                if docente not in docentes_asignados[fecha]:
                    return True
                
                # Verificar que no tenga otro examen a la misma hora
                return hora not in docentes_asignados[fecha][docente]
            
            # Función para verificar disponibilidad de grupo (solo para no optativas)
            def grupo_disponible(grupo, fecha, hora, es_optativa):
                if es_optativa:
                    return True  # Las optativas pueden coincidir en horario para el mismo grupo
                
                if fecha not in grupos_asignados:
                    return True
                if grupo not in grupos_asignados[fecha]:
                    return True
                
                # Verificar que el grupo no tenga otro examen a la misma hora (solo para no optativas)
                return hora not in grupos_asignados[fecha][grupo]
            
            # Generar horarios
            horarios = []
            
            # Primero asignar materias no optativas
            for semestre in todos_semestres:
                for grupo, materias_grupo in materias_por_grupo.get(semestre, {}).items():
                    # Separar materias por turno
                    materias_matutino = materias_grupo[materias_grupo['TURNO'] == 'M']
                    materias_vespertino = materias_grupo[materias_grupo['TURNO'] == 'V']
                    
                    # Asignar materias matutinas
                    for i in range(min(config["examenes_por_turno"], len(materias_matutino))):
                        materia = materias_matutino.iloc[i]
                        hora = config["horario_matutino"][i]
                        
                        # Buscar día disponible para este grupo y docente
                        for dia in dias_habiles:
                            if (docente_disponible(materia['DOCENTE'], dia, hora) and 
                                grupo_disponible(grupo, dia, hora, materia['ES_OPTATIVA'])):
                                
                                # Registrar asignación
                                if dia not in docentes_asignados:
                                    docentes_asignados[dia] = {}
                                if materia['DOCENTE'] not in docentes_asignados[dia]:
                                    docentes_asignados[dia][materia['DOCENTE']] = []
                                docentes_asignados[dia][materia['DOCENTE']].append(hora)
                                
                                if not materia['ES_OPTATIVA']:
                                    if dia not in grupos_asignados:
                                        grupos_asignados[dia] = {}
                                    if grupo not in grupos_asignados[dia]:
                                        grupos_asignados[dia][grupo] = []
                                    grupos_asignados[dia][grupo].append(hora)
                                
                                # Agregar al horario
                                horarios.append({
                                    "Fecha": dia,
                                    "Hora": hora,
                                    "Licenciatura": "LICENCIATURA EN CONTADURÍA",
                                    "Materia": materia['MATERIA'],
                                    "Docente": materia['DOCENTE'],
                                    "Semestre": materia['SEMESTRE'],
                                    "Grupo": materia['GRUPO'],
                                    "Turno": "Matutino",
                                    "EsOptativa": materia['ES_OPTATIVA']
                                })
                                break
                    
                    # Asignar materias vespertinas
                    for i in range(min(config["examenes_por_turno"], len(materias_vespertino))):
                        materia = materias_vespertino.iloc[i]
                        hora = config["horario_vespertino"][i]
                        
                        # Buscar día disponible para este grupo y docente
                        for dia in dias_habiles:
                            if (docente_disponible(materia['DOCENTE'], dia, hora) and 
                                grupo_disponible(grupo, dia, hora, materia['ES_OPTATIVA'])):
                                
                                # Registrar asignación
                                if dia not in docentes_asignados:
                                    docentes_asignados[dia] = {}
                                if materia['DOCENTE'] not in docentes_asignados[dia]:
                                    docentes_asignados[dia][materia['DOCENTE']] = []
                                docentes_asignados[dia][materia['DOCENTE']].append(hora)
                                
                                if not materia['ES_OPTATIVA']:
                                    if dia not in grupos_asignados:
                                        grupos_asignados[dia] = {}
                                    if grupo not in grupos_asignados[dia]:
                                        grupos_asignados[dia][grupo] = []
                                    grupos_asignados[dia][grupo].append(hora)
                                
                                # Agregar al horario
                                horarios.append({
                                    "Fecha": dia,
                                    "Hora": hora,
                                    "Licenciatura": "LICENCIATURA EN CONTADURÍA",
                                    "Materia": materia['MATERIA'],
                                    "Docente": materia['DOCENTE'],
                                    "Semestre": materia['SEMESTRE'],
                                    "Grupo": materia['GRUPO'],
                                    "Turno": "Vespertino",
                                    "EsOptativa": materia['ES_OPTATIVA']
                                })
                                break
            
            # Luego asignar materias optativas que no cupieron en la primera pasada
            for semestre in todos_semestres:
                for grupo, materias_grupo in materias_por_grupo.get(semestre, {}).items():
                    # Filtrar solo optativas no asignadas
                    optativas_no_asignadas = materias_grupo[
                        (materias_grupo['ES_OPTATIVA']) & 
                        (~materias_grupo['MATERIA'].isin([h['Materia'] for h in horarios if h['Grupo'] == grupo]))
                    ]
                    
                    for _, materia in optativas_no_asignadas.iterrows():
                        turno = "Matutino" if materia['TURNO'] == 'M' else "Vespertino"
                        horas_turno = config["horario_matutino"] if turno == "Matutino" else config["horario_vespertino"]
                        
                        for hora in horas_turno:
                            # Buscar día disponible para este docente
                            for dia in dias_habiles:
                                if docente_disponible(materia['DOCENTE'], dia, hora):
                                    # Registrar asignación
                                    if dia not in docentes_asignados:
                                        docentes_asignados[dia] = {}
                                    if materia['DOCENTE'] not in docentes_asignados[dia]:
                                        docentes_asignados[dia][materia['DOCENTE']] = []
                                    docentes_asignados[dia][materia['DOCENTE']].append(hora)
                                    
                                    # Agregar al horario
                                    horarios.append({
                                        "Fecha": dia,
                                        "Hora": hora,
                                        "Licenciatura": "LICENCIATURA EN CONTADURÍA",
                                        "Materia": materia['MATERIA'],
                                        "Docente": materia['DOCENTE'],
                                        "Semestre": materia['SEMESTRE'],
                                        "Grupo": materia['GRUPO'],
                                        "Turno": turno,
                                        "EsOptativa": True
                                    })
                                    break
                            else:
                                continue
                            break
            
            # Crear DataFrame final
            if not horarios:
                messagebox.showwarning("Advertencia", "No se generaron horarios. Verifique los datos de entrada.")
                return
            
            self.horarios_generados = pd.DataFrame(horarios)
            
            # Ordenar por fecha y hora
            self.horarios_generados['HoraOrden'] = pd.to_datetime(self.horarios_generados['Hora'], format='%H:%M')
            self.horarios_generados = self.horarios_generados.sort_values(['Fecha', 'HoraOrden', 'Semestre', 'Grupo'])
            self.horarios_generados = self.horarios_generados.drop('HoraOrden', axis=1)
            
            # Mostrar estadísticas
            stats = {
                'Total exámenes': len(horarios),
                'Días utilizados': len({h['Fecha'] for h in horarios}),
                'Semestres programados': ', '.join(sorted({h['Semestre'] for h in horarios})),
                'Grupos programados': ', '.join(sorted({h['Grupo'] for h in horarios})),
                'Materias optativas': sum(1 for h in horarios if h['EsOptativa'])
            }
            
            messagebox.showinfo(
                "Resultados",
                "Horarios generados:\n\n" + 
                "\n".join(f"{k}: {v}" for k, v in stats.items())
            )
            
            self.mostrar_resultados()
            self.export_btn.config(state=tk.NORMAL)
            
        except Exception as e:
            messagebox.showerror(
                "Error crítico",
                f"No se pudieron generar los horarios:\n\n{str(e)}\n\n"
                f"Detalles: {traceback.format_exc()}"
            )