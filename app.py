# ==========================================
    # ASIGNACIÓN INTELIGENTE DE DÍAS LIBRES (Basada en Demanda)
    # ==========================================
    for s in range(semanas_count):
        indices_semana = [s * 7 + i for i in range(7)]
        
        # Calcular la carga total de turnos requeridos por día en la semana actual
        carga_diaria = {idx: 0 for idx in indices_semana}
        for idx in indices_semana:
            fecha_temp = fechas_dt[idx]
            dia_nom = dias_semana_es[fecha_temp.weekday()]
            for c_k, v_dict in reglas_demanda.items():
                if "ANALISTA" not in c_k:
                    carga_diaria[idx] += len(v_dict.get(dia_nom, []))

        # Asignar libres a empleados que NO son analistas
        for _, emp in df_personal.iterrows():
            cod = str(emp["CODIGO"]).strip()
            cargo_emp = programacion_matriz[cod]["CARGO_ORIGINAL"]

            if "ANALISTA" not in cargo_emp:
                if cod not in dias_libres_emp:
                    dias_libres_emp[cod] = set()

                inicio_sem = fecha_base + timedelta(days=s * 7)
                fin_sem = inicio_sem + timedelta(days=6)
                hay_festivo_sem = any(inicio_sem.date() <= f <= fin_sem.date() for f in festivos_list)
                cant_libres = libres_base + (1 if hay_festivo_sem else 0)

                if cod in lideres:
                    idx_lider = lideres.index(cod)
                    dias_libres_emp[cod].add(s * 7 + (4 if idx_lider % 2 == 0 else 5))

                elif cod in tecnicos:
                    candidatos = [idx_d for idx_d in indices_semana if conteo_libres_tecnicos[idx_d] < 1]
                    elegidos = candidatos[:cant_libres] if len(candidatos) >= cant_libres else sorted(indices_semana, key=lambda x: conteo_libres_tecnicos[x])[:cant_libres]
                    for el in elegidos:
                        dias_libres_emp[cod].add(el)
                        conteo_libres_tecnicos[el] += 1
                else:
                    # AUXILIARES Y OTROS: Priorizar descansos en los días de MENOR carga operativa
                    dias_ordenados_por_carga = sorted(indices_semana, key=lambda idx: carga_diaria[idx])
                    
                    # Tomar los días con menor requerimiento para asignar el descanso
                    libres_elegidos = dias_ordenados_por_carga[:cant_libres]
                    dias_libres_emp[cod].update(libres_elegidos)

                    # Incrementar levemente la carga simulada para balancear los descansos entre el equipo
                    for el in libres_elegidos:
                        carga_diaria[el] += 2
