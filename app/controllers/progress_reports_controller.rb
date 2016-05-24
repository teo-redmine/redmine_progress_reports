# encoding: UTF-8
#
# ﻿Empresa desarrolladora: Fujitsu Technology Solutions S.A. - http://ts.fujitsu.com - Carlos Barroso Baltasar
#
# Autor: Junta de Andalucía
# Derechos de explotación propiedad de la Junta de Andalucía.
#
# Éste programa es software libre: usted tiene derecho a redistribuirlo y/o modificarlo bajo los términos de la Licencia EUPL European Public License publicada por el organismo IDABC de la Comisión Europea, en su versión 1.0. o posteriores.
#
# Éste programa se distribuye de buena fe, pero SIN NINGUNA GARANTÍA, incluso sin las presuntas garantías implícitas de USABILIDAD o ADECUACIÓN A PROPÓSITO CONCRETO. Para mas información consulte la Licencia EUPL European Public License.
#
# Usted recibe una copia de la Licencia EUPL European Public License junto con este programa, si por algún motivo no le es posible visualizarla, puede consultarla en la siguiente URL: http://ec.europa.eu/idabc/servlets/Docb4f4.pdf?id=31980
#
# You should have received a copy of the EUPL European Public License along with this program. If not, see
# http://ec.europa.eu/idabc/servlets/Docbb6d.pdf?id=31979
#
# Vous devez avoir reçu une copie de la EUPL European Public License avec ce programme. Si non, voir http://ec.europa.eu/idabc/servlets/Doc5a41.pdf?id=31983
#
# Sie erhalten eine Kopie der europäischen EUPL Public License zusammen mit diesem Programm. Wenn nicht, finden Sie da http://ec.europa.eu/idabc/servlets/Doc9dbe.pdf?id=31977

class ProgressReportsController < ApplicationController

  include ProgressReportsHelper
  include ActionView::Helpers::NumberHelper # Para el formato de números

  # Palabras clave del fichero ODS
  DIR_TECNICA = "$dir_tecnica"
  PRESUPUESTO = "$presupuesto"
  PROYECTO = "$proyecto"
  PERIODO = "$periodo"
  FECHA = "$fecha_informe"
  # Órdenes de trabajo
  ACTUACIONES="$actuaciones"
  ACT_PLANIFICACION="$act_planificacion"
  ACT_ESTIMADO="$act_estimado"
  ACT_RECURSOS="$act_recursos"
  ACT_DESVIACION="$act_desviacion"
  ACT_FINAL="$act_final"
  # Indicadores de nivel de servicio
  INDICADORES = "$indicadores"
  COSTES_IMPORTE = "$costes_importe"
  COSTES_PORCENTAJE = "$costes_porcentaje"

  # Costes de trabajo
  # Importe
  CT_IMPORTE_PRESUPUESTO = "$ct_importe_presupuesto"
  CT_IMPORTE_ASIGNADO = "$ct_importe_asignado"
  CT_IMPORTE_DESVIACION = "$ct_importe_desviacion"
  CT_IMPORTE_DISPONIBLE = "$ct_importe_disponible"
  CT_IMPORTE_CERRADO = "$ct_importe_cerrado"
  CT_IMPORTE_FACTURADO = "$ct_importe_facturado"
  CT_IMPORTE_TOTAL= "$ct_importe_total"
  # Porcentaje
  CT_PORCENTAJE_PRESUPUESTO = "$ct_porcentaje_presupuesto"
  CT_PORCENTAJE_ASIGNADO = "$ct_porcentaje_asignado"
  CT_PORCENTAJE_DESVIACION = "$ct_porcentaje_desviacion"
  CT_PORCENTAJE_DISPONIBLE = "$ct_porcentaje_disponible"
  CT_PORCENTAJE_CERRADO = "$ct_porcentaje_cerrado"
  CT_PORCENTAJE_FACTURADO = "$ct_porcentaje_facturado"
  CT_PORCENTAJE_TOTAL= "$ct_porcentaje_total"

  # Nombres de campos personalizados
  CF_IMPORTE_ESTIMADO = "Importe estimado"
  CF_IMPORTE_FINAL = "Importe final"
  CF_PRESUPUESTO = "Presupuesto"

  # Tipos de peticiones
  TRACKER_ACTUACION = "[AC] Actuación"
  TRACKER_OT = "[OT] Orden de Trabajo"
  TRACKER_OBJETIVO = "[OO] Objetivo Operativo"

  # Perfiles
  ROLE_DIR_TECNICA="Dirección Técnica"

  def generate
    # Variables globales #
    @issue = Issue.find(params[:issue_id]) # Petición Informe Progreso
    @project = Project.find(params[:project_id]) # Proyecto de la petición
    # Campo personalizado del proyecto: Presupuesto
    cv_budget = get_project_custom_value(CF_PRESUPUESTO)
    @budget =  if cv_budget
                cv_budget.value.to_f
              else
                0.to_f
              end
    # Modificamos el libro a partir de la plantilla
    @book = Rspreadsheet.open('plugins/redmine_progress_reports/report_template.ods')
    @sheet = @book.worksheets(1) # Informe de una única hoja
    # Campos costes de trabajos
    # Se van calculando mientras se insertan los datos de las OTs
    @wc_assigned = 0 # Importe Asignado
    @wc_estimated_deviation = 0 # Importe Desviación estimada
    @wc_closed = 0 # Importe Cerrado no facturado
    @wc_invoiced = 0 # Importe Facturado

    @act_estimated_available # Disponible estimado por actuación
    @objective_colis = Hash.new # Columnas de los objetivos

    # Peticiones para 'Recursos y contratos'
    @ots = Array.new

    begin
      ### Generación del informe ###

      insert_description # Sumario y previsiones
      insert_rows_objetivos(consulta_issues) # Objetivos
      insert_head_fields # Cabecera
      insert_performances # Actuaciones
      insert_working_costs_fields # Costes de trabajos
      insert_indicators # Indicadores de nivel de servicio
      insert_resources_contracts # Recursos y contratos

      ### Adjuntamos el informe ###
      report_name = @issue.to_s+".ods"
      # Obtenemos el contenido del informe en formato ODS
      report_content = get_report_ods_content(report_name)
      save_report_attachment(report_name,report_content) # Adjuntamos
    rescue Exception => e
      flash[:error] = l(:error_generate_report)
    end

    redirect_back_or_default issue_path(@issue) # Volvemos a visualizar la petición
  end

  private ##################################################################

  ### Campos de la cabecera ###

  def insert_head_fields
    # Miembros cuyo rol sea 'Dirección Técnica'
    directors = []
    @project.members.each do |m|
      director = m.roles.select { |r| r.name.eql?("Dirección Técnica") }.first
      directors << m.user.to_s if director
    end
    #Obtenemos las coordenadas de la celda Dir Técnica
    coord=find_cell_by_value(DIR_TECNICA).coordinates
    @sheet[coord[0],coord[1]]= directors.join(", ")

    #Obtenemos las coordenadas de la celda Proyecto
    coord=find_cell_by_value(PRESUPUESTO).coordinates
    @sheet[coord[0],coord[1]]= @budget
    @sheet.cell(coord[0],coord[1]).set_type_attribute(:currency)

    #Obtenemos las coordenadas de la celda Proyecto
    coord=find_cell_by_value(PROYECTO).coordinates

    #Insertamos en la celda Proyecto, el campo Nombre del proyecto de la petición
    @sheet[coord[0],coord[1]]= @project.name
  
    #Obtenemos las coordenadas de la celda Periodo
    coord=find_cell_by_value(PERIODO).coordinates

    #Insertamos en la celda Periodo, el Año de creación de la petición
    @sheet[coord[0],coord[1]]=@issue.created_on.year

    #Fecha de informe: Fecha de creación del informe(Fecha actual) - Celda que contiene la cadena '$fecha_informe'
    coord=find_cell_by_value(FECHA).coordinates
    @sheet[coord[0],coord[1]]=  Time.now.strftime("%d/%m/%Y").to_s # Solo dd/mm/yyyy
  end

  ### Adjuntar el informe ###

  def save_report_attachment(report_name,report_content)
	@issue.init_journal(User.current) # Nueva entrada en el histórico
    @attachment = Attachment.new(:file => report_content)
    @attachment.author = User.current
    @attachment.filename = report_name
    @attachment.content_type = @book.mime
    # Adjutamos el documento actualizando la petición
    if @attachment.save
      attachment={"1"=>{"filename"=>@attachment.filename, "description"=>"", "token"=>@attachment.token}}
      @issue.save_attachments(attachment)
      if @issue.save # Todo correcto
        flash[:notice] = l(:notice_successful_generate) unless @issue.current_journal.new_record?
      end
    end
    flash[:error] = l(:error_generate_report) if flash[:notice].empty? # Errores
  end

  def get_report_ods_content(report_name)
    @book.save('/tmp/'+report_name) # Guardamos el informe en la máquina temporalmente
    report = File.open('/tmp/'+report_name)
    report_content = IO.readlines(report).join("")
    File.delete(report) # Eliminamos el fichero de la máquina
    report_content
  end

  ### Objetivos ###

  def insert_rows_objetivos(peticiones = {})
    if peticiones.nil?
      #no hacer nada
    else
      #calculo rango de celdas
      rows = rows_between_cells("$ini_objetivos", "$fin_objetivos")
      #recojo columna
      start_rowi,start_coli = get_cell_rowi_coli_by("$ini_objetivos")
      #borro marcas de inicio y fin
      @sheet[rows.first.rowi, start_coli] = "" #borro marca de inicio
      @sheet[rows.last.rowi, start_coli] = "" #borro marca de fin
      #Por cada petición abro bucle (resto 1 para usar la columna que ya está definida en la plantilla)
      iteraciones = peticiones.length-1
      iteraciones.times do 
        rows.each_with_index do |i, x|
        #Inserto celda
          new_rowi = rows.first.rowi + x #calculo siguiente celda
          @sheet.insert_cell_before(new_rowi, start_coli)  # inserto celda
        end
      end
      #llamo al método para escribir mis objetivos.
      insert_objetivos(peticiones, rows, start_coli)
    end
  end

  def insert_objetivos(peticiones = {}, rows, start_coli)
    if peticiones.nil?
      #nada
    else
      peticiones.each_with_index do |issue, i|
        @sheet[rows.last.rowi, start_coli+i] = peticiones[i][:subject]
        # Guardamos las columnas
        @objective_colis.merge!({issue.id.to_s => start_coli+i})
      end
    end  
  end

  def consulta_issues
    @project.issues.select {|i| i.tracker.name.eql?(TRACKER_OBJETIVO)}
  end

  def find_cell_by_value(value)
    @sheet.nonemptycells.select {|nec| nec.value.eql?(value)}.first
  end

  ### Actuaciones ###

  # Insertar actuaciones
  def insert_performances

    # Celdas de referencia
    @act_rowi,@act_coli = get_cell_rowi_coli_by(ACTUACIONES) # Órdenes de trabajo
    @act_plan_coli = find_cell_by_value(ACT_PLANIFICACION).coli # Planificación
    @estimated_coli = find_cell_by_value(ACT_ESTIMADO).coli # Importe estimado
    @resources_coli = find_cell_by_value(ACT_RECURSOS).coli # % de recursos consumidos
    @deviation_coli = find_cell_by_value(ACT_DESVIACION).coli # Desviación estimada
    @final_coli = find_cell_by_value(ACT_FINAL).coli # Importe final

    # Eliminamos el contenido de las celdas de referencia
    delete_cell_value(@act_rowi,@act_plan_coli)
    delete_cell_value(@act_rowi,@estimated_coli)
    delete_cell_value(@act_rowi,@resources_coli)
    delete_cell_value(@act_rowi,@deviation_coli)
    delete_cell_value(@act_rowi,@final_coli)

    # Actuaciones del proyecto
    performances = @project.issues.select {|i| i.tracker.name.eql?(TRACKER_ACTUACION) && !i.closed?}

    last_index = performances.length - 1
    performances.each_with_index do |p, i|
      insert_performance_values(p)
      2.times { |i| @sheet.add_row_above(@act_rowi+=i) } if i !=  last_index
    end
  end

  # Valores de las actuaciones
  def insert_performance_values(performance)
    @sheet[@act_rowi,@act_coli] = performance.subject # Asunto Actuación
    delete_cell_value(@act_rowi, @estimated_coli)

    # Sólo OTs con fechas de inicio y fin rellenas
    children = performance.children.select { |child| 
      child.tracker.name.eql?(TRACKER_OT) && !child.start_date.to_s.empty? && !child.due_date.to_s.empty?
    }
    
    # Almacenamos las OTs y la actuación para el cálculo de imputaciones
    @ots += children 

    # Disponible estimado
    @act_estimated_available=0 # Inicializamos los sumatorios
    @act_total_deviation=0
    @sheet.add_row_above(@act_rowi) # Nueva fila

    # Importe estimado Actuación
    cv_estimated_amount = performance.custom_values.select{|cv| cv.custom_field.name.eql?(CF_PRESUPUESTO)}.first
    cv_estimated_amount = cv_estimated_amount.value if cv_estimated_amount
    @sheet[@act_rowi, @estimated_coli] = cv_estimated_amount.to_f
    @sheet.cell(@act_rowi, @estimated_coli).set_type_attribute(:currency)

    @sheet[@act_rowi+=1, @act_coli] = add_tabs_before(1, "Disponible estimado")
    delete_cell_value(@act_rowi, @estimated_coli)

    # OTs - En ejecución
    @sheet.add_row_above(@act_rowi) # Nueva fila
    @sheet[@act_rowi+=1, @act_coli] = add_tabs_before(1, "En ejecución")
    ots_in_progress = children.select { |child| ["En curso", "Resuelto"].include?(child.status.name) } 
    @act_rowi+=1

    if ots_in_progress.any?
      insert_ots(ots_in_progress)
      # Sumatorios
      sum_ot_totals(@act_rowi-ots_in_progress.length, @act_rowi-1, false)
    end

    # OTs - Cerradas
    ots_closed = children.select { |child| ["Cerrado", "Facturado"].include?(child.status.name) }
    # Inserta 2 filas
    2.times { |i| @sheet.add_row_above(@act_rowi+=i) }
    @sheet[@act_rowi, @act_coli] = add_tabs_before(1, "Cerrado")

    @act_rowi+=1
    if ots_closed.any?
      insert_ots(ots_closed)
      sum_ot_totals(@act_rowi-ots_closed.length, @act_rowi-1, true)
    end

    if (ots_in_progress.any? && ots_closed.any?)
      # Total 'Disponible estimado' : "En ejecución" + "Cerrado"
      # Quitamos 4 filas: vacía + "En ejecución"+ "Cerrado"+ siguiente
      total_estimated_rowi=@act_rowi-ots_closed.length-ots_in_progress.length-4
      @sheet[total_estimated_rowi,@estimated_coli] = @budget-(@act_estimated_available+@act_total_deviation)
      @sheet.cell(total_estimated_rowi, @estimated_coli).set_type_attribute(:currency)
    end
  end

  # Insertar las OT
  def insert_ots(ots)
    ots.each_with_index do |ot, i|
      @sheet.add_row_above(@act_rowi)
      @sheet[@act_rowi, @act_coli] = add_tabs_before(2, name(ot))

      insert_ot_related_objetives(ot)
      insert_ot_planification(ot)
      insert_ot_data(ot)
      @act_rowi+=1
    end
  end

  # OTs relacionadas con objetivos
  def insert_ot_related_objetives(ot)
    related_objetives = ot.relations.select {|r| r.issue_to.tracker.name.eql?(TRACKER_OBJETIVO) ||
                                                 r.issue_from.tracker.name.eql?(TRACKER_OBJETIVO)}
    related_objetives.each do |rot|
      coli =  if(rot.issue_to.eql?(@issue)) # to
                @objective_colis[rot.issue_to_id.to_s]
              else # From
                @objective_colis[rot.issue_from_id.to_s]
              end
      @sheet[@act_rowi, coli] = "X" if coli # Marcamos la relación
    end
  end

  def name(ot)
    ot.subject + " [#{ot.id}]"
  end

  # Planificación / Mes
  def insert_ot_planification(ot)
    dates_month_range = ot.start_date.month..ot.due_date.month
    te_months = ot.time_entries.collect {|te| te.spent_on.month }.uniq
    # Rellenamos los meses: del 0 al 11
    (0...12).each do |month|
      month_coli = month+@act_plan_coli
      if dates_month_range.include?(month) # Horas imputadas
        if te_months.include?(month)  # Si hay imputaciones en el mes
          @sheet.cell(@act_rowi, month_coli).value = "x"
        else # Dentro del rango sin horas imputadas
          @sheet[@act_rowi, month_coli] = "o"
        end         
      else # Fuera del rango de fechas
        @sheet[@act_rowi, month_coli] = "-"
      end 
    end
  end

  # Datos OT
  def insert_ot_data(ot)
    # 1 - Importe estimado: Campo personalizado
    estimated_amount = ot.custom_values.select{|cv| cv.custom_field.name.eql?(CF_IMPORTE_ESTIMADO)}.first
    if estimated_amount
      estimated_amount = estimated_amount.value.to_f
    end
    @sheet.cell(@act_rowi, @estimated_coli).value = estimated_amount
    @sheet.cell(@act_rowi, @estimated_coli).set_type_attribute('currency')
    # Almacenamos temporalmente este valor necesario en el cálculo 'Disponible estimado'
    @act_estimated_available += estimated_amount

    # 2 - % de recursos consumidos: Horas consumidas/estimadas
    percent_resources = if ot.estimated_hours.to_f > 0
                          ot.total_spent_hours.to_f/ot.estimated_hours.to_f
                        else
                          0
                        end
    @sheet.cell(@act_rowi,@resources_coli).value = percent_resources
    @sheet.cell(@act_rowi,@resources_coli).set_type_attribute('percentage')
    # 3 - Desviación estimada: ('% de recursos consumidos'-1)*'Importe estimado'
    if percent_resources > 1
      desviation = ((percent_resources - 1) * estimated_amount).round(2)
      @sheet.cell(@act_rowi,@deviation_coli).value = desviation
    else
      @sheet.cell(@act_rowi,@deviation_coli).value = 0.to_f
    end
    @sheet.cell(@act_rowi,@deviation_coli).set_type_attribute('currency')

    # 4 - Importe final: Campo personalizado
    final_amount = ot.custom_values.select{|cv| cv.custom_field.name.eql?(CF_IMPORTE_FINAL)}.first
    @sheet.cell(@act_rowi,@final_coli).value =  if final_amount
                                                  final_amount.value.to_f
                                                else
                                                  0.to_f
                                                end
    @sheet.cell(@act_rowi,@final_coli).set_type_attribute('currency')
  end

  # Suma totales OTs
  def sum_ot_totals(rowi_from, rowi_to,closed)
    row_range = rowi_from..rowi_to # Rango a sumar
    # 0 - Importe estimado
    # 1 - % de recursos consumidos
    # 2 - Desviación estimada
    # 3 - Importe final
    4.times do |n|
      coli = select_ot_values_column(n)
      total = if (n == 1) # % Recursos consumidos
                calculate_total_percent(row_range,coli)
              else # Resto
                sum_total_range(row_range,coli)
              end
      @sheet.cell(rowi_from-1,coli).value = total
      #Doy formato a la celda en función de si es porcentaje o moneda.
      if(n != 1)
        @sheet.cell(rowi_from-1,coli).set_type_attribute('currency')
      else
        @sheet.cell(rowi_from-1,coli).set_type_attribute('percentage')
      end
      @act_total_deviation += total if n == 2
      # Vamos sumando los totales asignados
      sum_working_costs_fields(n,total,closed)
    end
  end

  def select_ot_values_column(n)
    case n
    when 0 # Importe estimado
      @estimated_coli
    when 1 # % Recursos consumidos
      @resources_coli
    when 2 # Desviación estimada
      @deviation_coli
    when 3 #Importe final
      @final_coli 
    else
      nil
    end
  end

  def sum_working_costs_fields(n,total,closed)
    if (n==0) # Importe estimado
      @wc_assigned += total
      @wc_closed += total if closed
    elsif (n==2) # Desviación estimada
      @wc_estimated_deviation += total
    elsif (n==3) # Facturado
      @wc_invoiced += total
    end
  end

  def calculate_total_percent(row_range,coli)
    total = 0
    row_range.each do |r|
      # Multiplicamos cada porcentaje por su importe estimado
      total += @sheet.cell(r,coli).value.to_f * @sheet.cell(r,@estimated_coli).value.to_f
    end
    # Dividimos entre el total del Importe Estimado
    if @sheet.cell(row_range.first-1,@estimated_coli).value > 0
      (total/@sheet.cell(row_range.first-1,@estimated_coli).value).round(2)
    else
      0.to_f
    end
  end

  ### Indicadores de nivel de servicio ###

  def insert_indicators
    # Cabecera tabla
    indicators_rowi,indicators_coli = get_cell_rowi_coli_by("$indicadores")
    start_date_rowi,start_date_coli = get_cell_rowi_coli_by("$indicadores_fecha_inicio")
    end_date_rowi,end_date_coli = get_cell_rowi_coli_by("$indicadores_fecha_fin")
    done_rowi,done_coli = get_cell_rowi_coli_by("$indicadores_realizado")
    # Buscamos los objetivos
    objectives = get_project_open_issues_by_tracker("[OO] Objetivo Operativo")
    last_index = objectives.length-1
    # Insertamos los valores en las columnas
    objectives.each_with_index do |objetive,i|
      # Asunto
      @sheet.cell(indicators_rowi+i,indicators_coli).value = objetive.subject
      # Fecha de inicio: Solo dd/mm/yyyy
      @sheet.cell(start_date_rowi+i,start_date_coli).value = objetive.start_date.strftime("%d/%m/%Y").to_s
      # Fecha fin: Solo dd/mm/yyyy    
      @sheet.cell(end_date_rowi+i,end_date_coli).value = objetive.due_date.strftime("%d/%m/%Y").to_s
      # % Realizado
      @sheet.cell(done_rowi+i,done_coli).value = (objetive.done_ratio.to_f)/100
      # Nueva Fila
      @sheet.add_row_above(indicators_rowi+i) if i != last_index
    end
    # Establecemos el formato porcentaje
    (indicators_rowi..indicators_rowi+last_index).each do |rowi|
      @sheet.cell(rowi,done_coli).set_type_attribute('percentage')
    end
  end

  ### Costes de trabajos ###

  def insert_working_costs_fields
    # Campo 'Presupuesto' del proyecto
    insert_working_costs_fields_amounts # Importe
    insert_working_costs_fields_pertentages # % del presupuesto
  end

  # Importe
  def insert_working_costs_fields_amounts
    rowi, coli = get_cell_rowi_coli_by(CT_IMPORTE_PRESUPUESTO) # Celda de referencia
    @sheet.cell(rowi,coli).value=@budget # Presupuesto
    @sheet.cell(rowi,coli).set_type_attribute('currency') # Formato moneda
    @wc_estimated_available = @budget-(@wc_assigned+@wc_estimated_deviation) # Disponible estimado
    # Insertamos los valores
    6.times do |n| # 6 columnas
      field = select_working_costs_field_amount(n)
      cell_rowi,cell_coli = get_cell_rowi_coli_by(field)
      @sheet.cell(cell_rowi,cell_coli).value = select_working_cost(n)
      @sheet.cell(cell_rowi,cell_coli).set_type_attribute('currency') # Formato moneda
    end
  end

  # % del presupuesto
  def insert_working_costs_fields_pertentages
    rowi, coli = get_cell_rowi_coli_by(CT_PORCENTAJE_PRESUPUESTO) # Celda de referencia
    delete_cell_value(rowi,coli) # Eliminamos la palabra clave
    # Insertamos los valores
    6.times do |n| # 6 columnas
      field = select_working_costs_field_percentage(n)
      cell_rowi,cell_coli = get_cell_rowi_coli_by(field)
      @sheet.cell(cell_rowi,cell_coli).value = (select_working_cost(n)/@budget).round(2)
      @sheet.cell(cell_rowi,cell_coli).set_type_attribute('percentage')
    end
  end

  def select_working_costs_field_amount(n)
    case n
    when 0 # Asignado
      CT_IMPORTE_ASIGNADO
    when 1 # Desviación estimada
      CT_IMPORTE_DESVIACION
    when 2 # Disponible estimado
      CT_IMPORTE_DISPONIBLE
    when 3 # Cerrado no facturado
      CT_IMPORTE_CERRADO
    when 4 # Facturado
      CT_IMPORTE_FACTURADO
    when 5 # Total ejecutado
      CT_IMPORTE_TOTAL
    else
      ""
    end
  end

  def select_working_costs_field_percentage(n)
    case n
    when 0 # Asignado
      CT_PORCENTAJE_ASIGNADO
    when 1 # Desviación estimada
      CT_PORCENTAJE_DESVIACION
    when 2 # Disponible estimado
      CT_PORCENTAJE_DISPONIBLE
    when 3 # Cerrado no facturado
      CT_PORCENTAJE_CERRADO
    when 4 # Facturado
      CT_PORCENTAJE_FACTURADO
    when 5 # Total ejecutado
      CT_PORCENTAJE_TOTAL
    else
      ""
    end
  end

  def select_working_cost(n)
    case n
      when 0 # Asignado 
        @wc_assigned
      when 1 # Desviación estimada
        @wc_estimated_deviation
      when 2 # Disponible estimado
        @wc_estimated_available
      when 3 # Cerrado no facturado
        @wc_closed
      when 4 # Facturado
        @wc_invoiced
      when 5 # Total ejecutado
        @wc_invoiced+@wc_closed
      else
        nil
      end
  end

  ### Recursos y contratos ###

  def insert_resources_contracts
    # Celdas de referencia
    role_rowi,role_coli = get_cell_rowi_coli_by("$perfil") # Fila y columna
    amount_role_coli = find_cell_by_value("$importe_perfil").coli
    estimated_role_coli = find_cell_by_value("$previstas_perfil").coli
    spent_role_coli = find_cell_by_value("$consumidas_perfil").coli
    percentage_role_coli = find_cell_by_value("$porcentaje_perfil").coli
    total_coli = find_cell_by_value("$importe_total").coli

    # Cogemos todas las actividades menos "Horas base"
    horas_base_activity = Enumeration.where(type: "TimeEntryActivity", name: "Horas base")
    tes = TimeEntry.where(issue_id: @ots).where.not(activity_id: horas_base_activity) # Horas imputadas
    
    # Campos personalizados del proyecto: Coste y Horas previstas por perfil
    cf_costs = CustomField.where("name LIKE 'Coste %'")
    cf_estimated_hours = CustomField.where("name LIKE 'Horas prev. %'")

    project_cv_costs = @project.custom_values.select {|cv| cf_costs.include?(cv.custom_field)}
    project_cv_estimated_hours = @project.custom_values.select {|cv| cf_estimated_hours.include?(cv.custom_field) &&
                                                                      !cv.value.empty?}

    total = 0 # Total Importe previsto
    first_rowi = role_rowi
    last_index = project_cv_estimated_hours.length-1
    # Solo introducimos los perfiles que tengan relleno su campo de Horas previstas
    project_cv_estimated_hours.each_with_index do |eh,i|
      # Nombre del perfil
      role_short_name = eh.custom_field.name.sub("Horas prev. ","")
      role_name = select_role_name_by(role_short_name)
      if role_name.nil?
        role_name = role_short_name
      end
      # Horas consumidas
      tes.select { |te| te.activity.name.eql?(role_name) }
      role_spent_hours = tes.select {|te| te.activity.name.eql?(role_name)}.
                          inject(0){ |sum,te_role| sum += te_role.hours }

      # Importe estimado
      if role_name.include?("Asesor de microinformática")
        role_name.sub!("de ","")
      end
      cv_role_cost = project_cv_costs.select { |cost| cost.custom_field.name.include?(role_name) }.first
      role_cost = if cv_role_cost
                    cv_role_cost.value
                  else
                    0
                  end
      role_amount = role_cost.to_f * eh.value.to_f
      total+=role_amount
      # % unidades consumidas
      role_percentage_consumed = (role_spent_hours.to_f/eh.value.to_f).round(2)

      # Insertamos los valores
      @sheet.cell(role_rowi,role_coli).value = role_name.to_s # Nombre perfil
      @sheet.cell(role_rowi,amount_role_coli).value = role_amount # Importe previsto
      @sheet.cell(role_rowi,estimated_role_coli).value = eh.value.to_f # Unidades previstas
      @sheet.cell(role_rowi,spent_role_coli).value = role_spent_hours.to_f # Unidades consumidas
      @sheet.cell(role_rowi,percentage_role_coli).value = role_percentage_consumed # % Unidades consumidas
      #@sheet.cell(role_rowi,percentage_role_coli).set_type_attribute('percentage')
      # Nueva fila
      if i != last_index
        @sheet.add_row_above(role_rowi)
        role_rowi+=1 # Siguiente fila
      else
        @sheet.cell(role_rowi+1,total_coli).value = total # Total
      end
    end
    (first_rowi..role_rowi).each do |rowi|
      @sheet.cell(rowi,percentage_role_coli).set_type_attribute('percentage')
    end
  end

  def select_role_name_by(short_name)
    if short_name.include?("Analista program")
      "Analista programador"
    elsif short_name.include?("Técnico sistemas")
      "Técnico de sistemas"
    elsif short_name.include?("Analista sistemas")
      "Analista de sistemas"
    elsif short_name.include?("Administrador bbdd")
      "Administrador de base de datos"
    elsif short_name.include?("Asesor microinf")
      "Asesor de microinformática"
    else
      nil
    end
  end

  ### Sumario y previsiones ###

  def insert_description
    rowi,coli = get_cell_rowi_coli_by("$descripcion")
    @sheet.cell(rowi,coli).value = @issue.description
  end

  ### Funciones auxiliares ###

  # Suma los valores de una columna dentro de un rango de filas
  # Es posible eliminar el valor de cada una de las filas sumadas
  def sum_total_range(row_range,coli)
    sum = 0
    row_range.each do |r|
      sum += @sheet.cell(r,coli).value.to_f if @sheet.cell(r,coli).value
    end
    sum
  end

  # Elimina el contenido de una celda
  def delete_cell_value(rowi,coli)
    @sheet.cell(rowi,coli).value = ""
  end

  def get_project_custom_value(custom_field_name)
    @project.custom_values.select {|cv| cv.custom_field.name.eql?(custom_field_name)}.first
  end

  # Obtiene peticiones abiertas del proyecto de un tracker
  def get_project_open_issues_by_tracker(name)
    @project.issues.select {|i| i.tracker.name.eql?(name) && !i.closed?}
  end

  def get_cell_rowi_coli_by(value)
    cell = find_cell_by_value(value)
    [cell.rowi,cell.coli]
  end

  def row_range_between_cells(from, to)
    cell_from = find_cell_by_value(from)
    cell_to = find_cell_by_value(to)
    cell_from.row.rowi..cell_to.row.rowi if cell_from && cell_to
  end

  def rows_between_cells(from, to)    
    row_range = row_range_between_cells(from, to)
    row_range.collect {|rowi| @sheet.rows(rowi)}.flatten
  end

end
