else:
                # 1. Si es día libre asignado
                if idx_dia in dias_libres_emp.get(cod_emp, set()):
                    programacion_matriz[cod_emp][col_nombre] = "L"
                    d_emp["TURNO_FIJO_BLOQUE"] = None
                    d_emp["SALIDA_PREVIA_DT"] = None
                    d_emp["ULTIMA_FRANJA"] = None
                    d_emp["HISTORIAL_CARGOS_DIARIOS"][col_nombre] = cargo_orig
                    d_emp["HISTORIAL_TURNOS_LIMPIOS"][col_nombre] = "L"
                    continue

                turno_a_asignar = None
                cargo_cubierto_hoy = cargo_orig
                cargos_a_evaluar = [cargo_orig, cargo_sec] if cargo_sec else [cargo_orig]

                for c_eval in cargos_a_evaluar:
                    if not c_eval:
                        continue
                    
                    turnos_disp_cargo = demandas_dia_actual.get(c_eval, [])
                    if not turnos_disp_cargo:
                        continue

                    # Búsqueda 1: Intentar mantener bloque si existe y está en demanda
                    cand_bloque = d_emp.get("TURNO_FIJO_BLOQUE")
                    if cand_bloque and cand_bloque in turnos_disp_cargo:
                        parsed_h = extraer_horas(cand_bloque)
                        if parsed_h:
                            h_i, m_i, _, _ = parsed_h
                            entrada_dt = fecha_col.replace(hour=h_i, minute=m_i)
                            if calcular_descanso_suficiente(d_emp["SALIDA_PREVIA_DT"], entrada_dt, min_horas=12):
                                turno_a_asignar = cand_bloque
                                turnos_disp_cargo.remove(cand_bloque)
                                cargo_cubierto_hoy = c_eval
                                break

                    # Búsqueda 2: Si no funcionó el bloque, iterar sobre TODOS los turnos requeridos disponibles
                    franja_ult = d_emp["ULTIMA_FRANJA"]
                    franjas_permitidas = obtener_siguiente_franja_permitida(franja_ult) if franja_ult else ["NOCHE", "TARDE", "MAÑANA"]
                    
                    # Ordenar por franja permitida
                    cand_list = sorted(list(turnos_disp_cargo), key=lambda t: 0 if clasificar_franja(t) in franjas_permitidas else 1)

                    for cand_t in cand_list:
                        parsed_h = extraer_horas(cand_t)
                        if parsed_h:
                            h_i, m_i, _, _ = parsed_h
                            entrada_dt = fecha_col.replace(hour=h_i, minute=m_i)
                            if calcular_descanso_suficiente(d_emp["SALIDA_PREVIA_DT"], entrada_dt, min_horas=12):
                                turno_a_asignar = cand_t
                                turnos_disp_cargo.remove(cand_t)
                                cargo_cubierto_hoy = c_eval
                                break
                    
                    if turno_a_asignar:
                        break

                # Asignación final del resultado
                if turno_a_asignar is None:
                    # Búsqueda de emergencia: Ignorar preferencia de franja pero respetar el descanso legal
                    for c_eval in cargos_a_evaluar:
                        turnos_disp_cargo = demandas_dia_actual.get(c_eval, [])
                        for cand_t in list(turnos_disp_cargo):
                            parsed_h = extraer_horas(cand_t)
                            if parsed_h:
                                h_i, m_i, _, _ = parsed_h
                                entrada_dt = fecha_col.replace(hour=h_i, minute=m_i)
                                if calcular_descanso_suficiente(d_emp["SALIDA_PREVIA_DT"], entrada_dt, min_horas=12):
                                    turno_a_asignar = cand_t
                                    turnos_disp_cargo.remove(cand_t)
                                    cargo_cubierto_hoy = c_eval
                                    break
                        if turno_a_asignar:
                            break

                if turno_a_asignar is None:
                    programacion_matriz[cod_emp][col_nombre] = "AO"
                    d_emp["TURNO_FIJO_BLOQUE"] = None
                    d_emp["SALIDA_PREVIA_DT"] = None
                    d_emp["HISTORIAL_CARGOS_DIARIOS"][col_nombre] = cargo_orig
                    d_emp["HISTORIAL_TURNOS_LIMPIOS"][col_nombre] = "AO"
                else:
                    d_emp["TURNO_FIJO_BLOQUE"] = turno_a_asignar
                    d_emp["ULTIMA_FRANJA"] = clasificar_franja(turno_a_asignar)
                    d_emp["HISTORIAL_CARGOS_DIARIOS"][col_nombre] = cargo_cubierto_hoy
                    d_emp["HISTORIAL_TURNOS_LIMPIOS"][col_nombre] = turno_a_asignar

                    if cargo_cubierto_hoy != cargo_orig:
                        sufijo = obtener_iniciales_cargo(cargo_cubierto_hoy)
                        programacion_matriz[cod_emp][col_nombre] = f"{turno_a_asignar} {sufijo}"
                    else:
                        programacion_matriz[cod_emp][col_nombre] = turno_a_asignar

                    parsed_h = extraer_horas(turno_a_asignar)
                    if parsed_h:
                        h_i, m_i, h_f, m_f = parsed_h
                        if h_f >= 24:
                            d_emp["SALIDA_PREVIA_DT"] = (fecha_col + timedelta(days=1)).replace(hour=h_f - 24, minute=m_f)
                        elif h_f < h_i:
                            d_emp["SALIDA_PREVIA_DT"] = (fecha_col + timedelta(days=1)).replace(hour=h_f, minute=m_f)
                        else:
                            d_emp["SALIDA_PREVIA_DT"] = fecha_col.replace(hour=h_f, minute=m_f)
