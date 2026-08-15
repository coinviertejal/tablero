from pathlib import Path
from io import BytesIO
import hmac, pandas as pd, plotly.express as px, plotly.graph_objects as go, streamlit as st
from src.importers import parse_alzex,load_budget,load_simple_budget,load_extraordinary,load_compiled_monthly
from src.storage import client,fetch,insert_one,insert_rows
ROOT=Path(__file__).parent;DATA=ROOT/'data'/'initial';MONTHS={1:'Enero',2:'Febrero',3:'Marzo',4:'Abril',5:'Mayo',6:'Junio',7:'Julio',8:'Agosto',9:'Septiembre',10:'Octubre',11:'Noviembre',12:'Diciembre'};MONTH_NUM={v:k for k,v in MONTHS.items()};NAVY='#172A46';BLUE='#2563EB';SKY='#60A5FA';GOLD='#D59A33';RED='#DC2626';GREEN='#16A34A';GRID='#E5EAF1';MONTH_COLORS=['#2563EB','#F59E0B','#10B981','#8B5CF6','#EF4444','#06B6D4','#F97316','#6366F1','#84CC16','#EC4899','#14B8A6','#64748B']
APP_VERSION='2026.08.14-presupuesto-v2-graficas-individuales'
st.set_page_config(page_title='Presupuesto Familiar',page_icon='💰',layout='wide')
PLOT_CONFIG={'displaylogo':False,'responsive':True,'scrollZoom':True,'toImageButtonOptions':{'format':'png','filename':'presupuesto-familiar','scale':2}}
st.markdown("""<style>.stApp{background:#F7F9FC}.block-container{padding-top:2rem;max-width:1500px}h1,h2,h3{color:#172A46!important}.stMetric{background:white;border:1px solid #E5EAF1;border-radius:14px;padding:16px;box-shadow:0 2px 8px #172A4610}[data-testid='stSidebar']{background:#172A46}[data-testid='stSidebar'] *{color:#F8FAFC!important}.stDataFrame{border:1px solid #E5EAF1;border-radius:12px;overflow:hidden}</style>""",unsafe_allow_html=True)
def authenticate():
    expected=st.secrets.get('APP_PASSWORD','')
    if not expected:st.error('Falta configurar APP_PASSWORD.');st.stop()
    if st.session_state.get('authenticated'):return
    _,col,_=st.columns([1,1.1,1])
    with col:
        st.markdown('<div style="height:12vh"></div>',unsafe_allow_html=True);st.title('Presupuesto Familiar')
        with st.form('login'):pwd=st.text_input('Contraseña',type='password');sent=st.form_submit_button('Entrar',use_container_width=True,type='primary')
        if sent:
            if hmac.compare_digest(pwd,expected):st.session_state.authenticated=True;st.rerun()
            st.error('Contraseña incorrecta')
    st.stop()
@st.cache_data
def budget_data():return load_budget(str(DATA/'ppto_2026_morus.xlsx'))
@st.cache_data
def budget_template():
    sample=pd.DataFrame({'Categoría':['Hogar','Salud','Transporte','Compras personales','Ciencias'],'Monto':[35000,12000,9000,6000,10000]});buffer=BytesIO()
    with pd.ExcelWriter(buffer,engine='openpyxl') as writer:sample.to_excel(writer,index=False,sheet_name='Presupuesto mensual')
    return buffer.getvalue()
@st.cache_data
def extraordinary_data():return load_extraordinary(str(DATA/'gastos_no_programados.xlsx'))
def initial_movements():
    p=DATA/'alzex_julio_2026.csv';return parse_alzex(p.read_bytes(),p.name)
def compiled_data():return load_compiled_monthly(str(DATA/'2026_ene_jul.xlsx'))
def money(v):return f'${v:,.2f}' if abs(v-round(v))>.001 else f'${v:,.0f}'
def change_color(v):
    try:n=float(str(v).replace('$','').replace(',',''))
    except (TypeError,ValueError):return ''
    return 'color:#15803D;font-weight:700' if n>0 else ('color:#DC2626;font-weight:700' if n<0 else '')
def style(fig):
    fig.update_layout(font=dict(color=NAVY),paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)',margin=dict(l=10,r=10,t=52,b=10),legend_title_text='');fig.update_xaxes(gridcolor=GRID);fig.update_yaxes(gridcolor=GRID);return fig
def db_movements():
    rows=fetch('movements');df=pd.DataFrame(rows) if rows else initial_movements()
    if not df.empty:df['movement_date']=pd.to_datetime(df.movement_date);df['amount']=pd.to_numeric(df.amount)
    return df
def analytical_monthly():
    base=compiled_data().copy();mov=db_movements()
    if not mov.empty:
        mov['subcategory']=mov.get('subcategory','Sin detalle').fillna('').replace('','Sin detalle');tx=mov.assign(year=mov.movement_date.dt.year,month=mov.movement_date.dt.month,month_name=mov.movement_date.dt.month.map(MONTHS)).groupby(['year','month','month_name','category','subcategory'],as_index=False).amount.sum();tx['source']='Alzex';keys=set(zip(tx.year,tx.month));base=base[~base[['year','month']].apply(tuple,axis=1).isin(keys)];base=pd.concat([base,tx],ignore_index=True)
    return base
def period_filter(df,key):
    years=sorted(df.year.unique(),reverse=True);c1,c2=st.columns([1,3]);year=c1.selectbox('Año',years,key='y'+key);available=sorted(df.loc[df.year==year,'month'].unique());names=[MONTHS[m] for m in available];chosen=c2.multiselect('Meses a analizar',names,default=names,key='months'+key);selected=sorted(MONTH_NUM[m] for m in chosen)
    if not selected:st.info('Selecciona al menos un mes.');st.stop()
    scoped=df[(df.year==year)&(df.month.isin(selected))].copy();categories=sorted(scoped.category.dropna().unique());chosen_cat=st.multiselect('Categorías (opcional)',categories,key='cats'+key,placeholder='Todas las categorías')
    if chosen_cat:scoped=scoped[scoped.category.isin(chosen_cat)]
    subcats=sorted(scoped.subcategory.fillna('Sin detalle').unique());chosen_sub=st.multiselect('Subcategorías (opcional)',subcats,key='subs'+key,placeholder='Todas las subcategorías')
    if chosen_sub:
        scoped=scoped[scoped.subcategory.fillna('Sin detalle').isin(chosen_sub)];present=set(scoped.month.unique());missing=[MONTHS[m] for m in selected if m not in present]
        if missing:st.caption('Sin movimientos para este filtro: '+', '.join(missing)+'. Se mostrarán con valor $0 en la gráfica mensual.')
    label=f"{', '.join(chosen)} {year}" if len(chosen)<=3 else f'{len(chosen)} meses de {year}'
    return scoped,selected,label,year

def extraordinary_all():
    base=extraordinary_data().copy();base['month']=base.month.astype(str).str.strip().str.title();base['month_num']=base.month.map(MONTH_NUM);base['year']=2026;rows=fetch('extraordinary_expenses')
    if rows:
        saved=pd.DataFrame(rows);saved['expense_date']=pd.to_datetime(saved.expense_date);saved=saved.assign(month=saved.expense_date.dt.month.map(MONTHS),month_num=saved.expense_date.dt.month,year=saved.expense_date.dt.year).rename(columns={'concept':'concept','amount':'amount'});base=pd.concat([base[['month','concept','amount','month_num','year']],saved[['month','concept','amount','month_num','year']]],ignore_index=True)
    base['amount']=pd.to_numeric(base.amount,errors='coerce').fillna(0);return base

def horizontal_month_chart(view,title):
    data=view.groupby(['category','month','month_name'],as_index=False).amount.sum();order=data.groupby('category').amount.sum().sort_values(ascending=False).head(12);data=data[data.category.isin(order.index)];fig=px.bar(data,x='amount',y='category',color='month_name',orientation='h',barmode='stack',title=title,category_orders={'category':list(reversed(order.index)),'month_name':[MONTHS[m] for m in sorted(data.month.unique())]},color_discrete_sequence=MONTH_COLORS,labels={'amount':'Gasto','category':'Categoría','month_name':'Mes'});fig.update_xaxes(tickprefix='$',tickformat=',.0f');totals=data.groupby('category',as_index=False).amount.sum();fig.add_scatter(x=totals.amount,y=totals.category,mode='text',text=[money(v) for v in totals.amount],textposition='middle right',showlegend=False,hoverinfo='skip');fig=style(fig);fig.update_layout(margin=dict(l=10,r=95,t=52,b=10));return fig

def vertical_composition(view,y_step=None):
    data=view.groupby(['category','month','month_name'],as_index=False).amount.sum();order=data.groupby('category').amount.sum().sort_values(ascending=False).index;data=data[data.category.isin(order)];fig=px.bar(data,x='category',y='amount',color='month_name',barmode='group',title='Composición del gasto por mes',category_orders={'category':list(order),'month_name':[MONTHS[m] for m in sorted(data.month.unique())]},color_discrete_sequence=MONTH_COLORS,labels={'amount':'Gasto','category':'Categoría','month_name':'Mes'});fig.update_yaxes(tickprefix='$',tickformat=',.0f',dtick=y_step);fig.update_xaxes(tickangle=-35);return style(fig)

def all_categories_chart(view):
    data=view.groupby('category',as_index=False).amount.sum().sort_values('amount');fig=px.bar(data,x='amount',y='category',orientation='h',title='Gasto total por todas las categorías',text=[money(v) for v in data.amount],color='amount',color_continuous_scale=['#BFDBFE','#2563EB'],labels={'amount':'Gasto','category':'Categoría'});fig.update_xaxes(tickprefix='$',tickformat=',.0f');fig.update_traces(textposition='outside');fig.update_layout(coloraxis_showscale=False,margin=dict(l=10,r=90,t=52,b=10),height=max(450,36*len(data)));return style(fig)
authenticate()
with st.sidebar:
    st.markdown('## 💰 Presupuesto');page=st.radio('Navegación',['Resumen','Tendencias y fugas','Presupuesto','Extraordinarios','Inversiones','Importar Alzex'],label_visibility='collapsed');st.divider()
    with st.expander('Cargar presupuesto mensual'):
        st.download_button('Descargar plantilla Excel',budget_template(),file_name='plantilla_presupuesto_mensual.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',use_container_width=True)
        upload_key=st.session_state.get('budget_upload_key',0);budget_file=st.file_uploader('Excel con Categoría y Monto',type=['xlsx'],key=f'budget_file_{upload_key}')
        if budget_file:
            try:
                custom_budget=load_simple_budget(budget_file.getvalue());st.session_state['custom_budget']=custom_budget.to_dict('records');st.success(f'{len(custom_budget)} categorías cargadas.')
            except Exception as e:st.error(str(e))
        if st.session_state.get('custom_budget') and st.button('Restaurar presupuesto original',use_container_width=True):
            st.session_state.pop('custom_budget',None);st.session_state['budget_upload_key']=upload_key+1;st.rerun()
        if st.session_state.get('custom_budget'):st.caption('Fuente activa: archivo cargado')
        else:st.caption('Fuente activa: presupuesto original')
    st.caption('🔒 Datos privados');st.caption('Versión '+APP_VERSION)
    if st.button('Cerrar sesión',use_container_width=True):st.session_state.clear();st.rerun()
monthly=analytical_monthly();budget=pd.DataFrame(st.session_state['custom_budget']) if st.session_state.get('custom_budget') else budget_data();extra=extraordinary_all();monthly_budget=float(budget.monthly_budget.sum())
if page=='Resumen':
    st.title('Resumen financiero');st.caption('Dónde estás parado, qué se está desviando y dónde hay fugas');view,selected,label,year=period_filter(monthly,'sum');spent=float(view.amount.sum());target=monthly_budget*len(selected);delta=spent-target;extra_period=extra[(extra.year==year)&(extra.month_num.isin(selected))];st.markdown('#### '+label);a,b,c,d=st.columns(4);a.metric('Gasto registrado',money(spent),f'{money(delta)} vs. presupuesto',delta_color='inverse');b.metric('Presupuesto del periodo',money(target));c.metric('Extraordinarios del periodo',money(float(extra_period.amount.sum())));d.metric('Promedio mensual',money(spent/max(1,len(selected))));(st.warning if delta>0 else st.success)(f"El gasto está {money(abs(delta))} {'arriba' if delta>0 else 'debajo'} del presupuesto.")
    left,right=st.columns([1.35,1])
    with left:
        trend=view.groupby('month',as_index=False).amount.sum().set_index('month').reindex(selected,fill_value=0).rename_axis('month').reset_index();trend['month_name']=trend.month.map(MONTHS);fig=px.bar(trend,x='month_name',y='amount',title='Gasto mensual',text_auto=',.0f',color='month_name',color_discrete_sequence=MONTH_COLORS,labels={'month_name':'Mes','amount':'Gasto'});fig.add_hline(y=monthly_budget,line_dash='dash',line_color=GOLD,annotation_text=f'Presupuesto {money(monthly_budget)}');fig.update_yaxes(tickprefix='$',tickformat=',.0f');fig.update_layout(showlegend=False);st.plotly_chart(style(fig),use_container_width=True,config=PLOT_CONFIG)
    with right:
        st.plotly_chart(horizontal_month_chart(view,'Principales categorías por mes'),use_container_width=True,config=PLOT_CONFIG)
elif page=='Tendencias y fugas':
    st.title('Tendencias y fugas');st.caption('Evolución, concentración y categorías que más presionan el gasto');view,selected,label,year=period_filter(monthly,'trend');st.markdown('#### '+label);axis_choice=st.selectbox('Detalle del eje Y',['Automático','$5,000','$1,000'],key='axis_detail');y_step={'Automático':None,'$5,000':5000,'$1,000':1000}[axis_choice];st.caption('También puedes acercar una zona arrastrando sobre la gráfica y descargarla con el icono de cámara.');trend=view.groupby('month',as_index=False).amount.sum().set_index('month').reindex(selected,fill_value=0).rename_axis('month').reset_index();trend['month_name']=trend.month.map(MONTHS);trend['media']=trend.amount.rolling(3,min_periods=1).mean();fig=go.Figure();fig.add_bar(x=trend.month_name,y=trend.amount,name='Gasto mensual',marker_color=[MONTH_COLORS[m-1] for m in trend.month],text=[money(x) for x in trend.amount],textposition='outside');fig.add_scatter(x=trend.month_name,y=trend.media,name='Promedio móvil 3 meses',line=dict(color=NAVY,width=3),mode='lines+markers');fig.add_hline(y=monthly_budget,line_dash='dash',line_color=GOLD,annotation_text='Presupuesto mensual');fig.update_yaxes(tickprefix='$',tickformat=',.0f',dtick=y_step);fig.update_layout(title='Gasto y tendencia mensual');st.plotly_chart(style(fig),use_container_width=True,config=PLOT_CONFIG);st.plotly_chart(vertical_composition(view,y_step),use_container_width=True,config=PLOT_CONFIG);current=view.groupby('category',as_index=False).amount.sum().sort_values('amount',ascending=False);current['share']=current.amount/current.amount.sum();current.columns=['Categoría','Gasto','Participación'];current['Gasto']=current.Gasto.map(money);current['Participación']=current.Participación.map(lambda x:f'{x:.1%}');st.subheader('Concentración del gasto');st.dataframe(current,hide_index=True,use_container_width=True,column_config={'Categoría':st.column_config.TextColumn(width='large'),'Gasto':st.column_config.TextColumn(width='medium'),'Participación':st.column_config.TextColumn(width='small')});st.plotly_chart(all_categories_chart(view),use_container_width=True,config=PLOT_CONFIG)
elif page=='Presupuesto':
    st.title('Presupuesto contra gasto real');st.caption('Comparación mensual o acumulada con categorías conciliadas');view,selected,label,year=period_filter(monthly,'bud');plan=budget.groupby('category').monthly_budget.sum()*len(selected);actual=view.groupby('category').amount.sum();comp=pd.concat([plan.rename('Presupuesto'),actual.rename('Real')],axis=1).fillna(0);comp['Variación']=comp.Real-comp.Presupuesto;comp['Ejercicio']=comp.Real/comp.Presupuesto.replace(0,pd.NA);comp=comp.reset_index().rename(columns={'category':'Categoría'}).sort_values('Variación',ascending=False);st.markdown('#### '+label);c1,c2,c3=st.columns(3);c1.metric('Presupuesto',money(comp.Presupuesto.sum()));c2.metric('Gasto real',money(comp.Real.sum()));c3.metric('Variación',money(comp.Variación.sum()),delta_color='inverse');display=comp.copy();display['Presupuesto']=display.Presupuesto.map(money);display['Real']=display.Real.map(money);display['Variación']=display.Variación.map(money);display['Ejercicio']=display.Ejercicio.map(lambda x:'—' if pd.isna(x) else f'{x:.1%}');styled=display.style.map(lambda v:f'color:{GREEN};font-weight:600' if str(v).startswith('$-') else f'color:{RED};font-weight:600',subset=['Variación']);st.dataframe(styled,hide_index=True,use_container_width=True)
elif page=='Extraordinarios':
    st.title('Gastos extraordinarios');years=sorted(extra.year.dropna().unique(),reverse=True);c1,c2=st.columns([1,3]);eyear=c1.selectbox('Año',years,key='extra_year');enames=[MONTHS[m] for m in sorted(extra.loc[extra.year==eyear,'month_num'].dropna().unique())];echosen=c2.multiselect('Meses a analizar',enames,default=enames,key='extra_months');emonths=[MONTH_NUM[m] for m in echosen];extra_view=extra[(extra.year==eyear)&(extra.month_num.isin(emonths))];st.metric('Total del periodo seleccionado',money(float(extra_view.amount.sum())))
    with st.expander('Agregar gasto extraordinario'):
        with st.form('extraordinary'):
            f1,f2,f3=st.columns([1,2,1]);date=f1.date_input('Fecha');concept=f2.text_input('Concepto');amount=f3.number_input('Importe',min_value=0.0,step=100.0);sent=st.form_submit_button('Agregar renglón',type='primary')
        if sent:
            if not concept.strip() or amount<=0:st.error('Escribe concepto e importe.')
            else:
                try:insert_one('extraordinary_expenses',{'expense_date':date.isoformat(),'concept':concept.strip(),'amount':amount});st.rerun()
                except Exception as e:st.error(str(e))
    show=extra_view.rename(columns={'month':'Mes','concept':'Concepto','amount':'Importe'})[['Mes','Concepto','Importe']];show['Importe']=show.Importe.map(money);st.dataframe(show.style.set_properties(subset=['Importe'],**{'text-align':'center'}),hide_index=True,use_container_width=True);chart=extra_view.groupby(['month_num','month'],as_index=False).amount.sum().sort_values('month_num');fig=px.bar(chart,x='month',y='amount',title='Gastos extraordinarios por mes',text_auto=',.0f',color='month',color_discrete_sequence=MONTH_COLORS,labels={'month':'Mes','amount':'Importe'});fig.update_yaxes(tickprefix='$',tickformat=',.0f');fig.update_traces(texttemplate='$%{y:,.0f}',textposition='outside');st.plotly_chart(style(fig),use_container_width=True,config=PLOT_CONFIG)
elif page=='Inversiones':
    st.title('Inversiones y rendimientos');st.caption('Registra cada inversión y actualiza su valuación cuando quieras');investments=pd.DataFrame(fetch('investments'));valuations=pd.DataFrame(fetch('investment_valuations'))
    if not investments.empty:
        investments['balance']=pd.to_numeric(investments['balance']);investments['label']=investments['institution'].astype('string').fillna('')+' · '+investments['product'].astype('string').fillna('');inflation,cetes=st.columns(2);inflation_rate=inflation.number_input('Inflación anual de referencia %',min_value=0.0,value=4.0,step=0.1)/100;cetes_rate=cetes.number_input('CETES anual de referencia %',min_value=0.0,value=8.0,step=0.1)/100
        if not valuations.empty:
            valuations['value']=pd.to_numeric(valuations.value);valuations['valuation_date']=pd.to_datetime(valuations.valuation_date);valuations=valuations.sort_values(['investment_id','valuation_date'])
        summary=[]
        for _,inv in investments.iterrows():
            history=valuations[valuations['investment_id']==inv['id']].copy() if not valuations.empty else pd.DataFrame()
            if history.empty:first=latest=float(inv['balance']);previous=pd.NA;days=0
            else:first=float(history.iloc[0]['value']);latest=float(history.iloc[-1]['value']);previous=float(history.iloc[-2]['value']) if len(history)>1 else pd.NA;days=max(0,(history.iloc[-1]['valuation_date']-history.iloc[0]['valuation_date']).days)
            abs_change=latest-previous if pd.notna(previous) else pd.NA;pct_change=(latest/previous-1) if pd.notna(previous) and previous else pd.NA;total_return=(latest/first-1) if len(history)>1 and first else pd.NA;annualized=((latest/first)**(365/days)-1) if len(history)>1 and days>0 and first>0 and latest>=0 else pd.NA
            projected=latest*(1+annualized) if pd.notna(annualized) else pd.NA;inflation_value=latest*(1+inflation_rate);cetes_value=latest*(1+cetes_rate);summary.append({'Inversión':inv['label'],'Valuaciones':len(history),'Valor actual':latest,'Cambio último':abs_change,'% último':pct_change,'Rendimiento total':total_return,'Proyección anual':annualized,'Proyección 12m':projected,'Referencia inflación 12m':inflation_value,'Diferencia vs. inflación':projected-inflation_value if pd.notna(projected) else pd.NA,'Referencia CETES 12m':cetes_value,'Diferencia vs. CETES':projected-cetes_value if pd.notna(projected) else pd.NA})
        summary=pd.DataFrame(summary);st.metric('Patrimonio invertido',money(summary['Valor actual'].sum()));display=summary.copy()
        for col in ['Valor actual','Cambio último','Proyección 12m','Referencia inflación 12m','Diferencia vs. inflación','Referencia CETES 12m','Diferencia vs. CETES']:display[col]=display[col].map(lambda x:'—' if pd.isna(x) else money(x))
        for col in ['% último','Rendimiento total','Proyección anual']:display[col]=display[col].map(lambda x:'—' if pd.isna(x) else f'{x:.2%}')
        st.dataframe(display.style.map(change_color,subset=['Cambio último']),hide_index=True,use_container_width=True);st.caption('“Cambio último” y “% último” comparan las dos valuaciones más recientes registradas para cada inversión.')
        with st.expander('Registrar nueva valuación'):
            valuation_investments={f"{row['label']} · ID {int(row['id'])}":int(row['id']) for _,row in investments.iterrows()}
            with st.form('valuation'):
                v1,v2=st.columns(2);selected_label=v1.selectbox('Inversión',list(valuation_investments));vdate=v2.date_input('Fecha de valuación');value=v1.number_input('Valor actual',min_value=0.0,step=100.0);notes=v2.text_input('Nota opcional');save_value=st.form_submit_button('Guardar valuación',type='primary')
            if save_value:
                inv_id=valuation_investments[selected_label]
                try:insert_one('investment_valuations',{'investment_id':inv_id,'valuation_date':vdate.isoformat(),'value':value,'notes':notes});st.success('Valuación registrada');st.rerun()
                except Exception as e:st.error('No se pudo guardar. Si ya existe una valuación de ese día, usa otra fecha. '+str(e))
        if not valuations.empty:
            with st.expander('Consultar y corregir histórico de valuaciones'):
                history_options={f"{row['label']} · ID {int(row['id'])}":int(row['id']) for _,row in investments.iterrows()}
                history_label=st.selectbox('Consultar inversión',list(history_options),key='history_investment');history_id=history_options[history_label]
                selected_history=valuations[valuations['investment_id']==history_id].sort_values('valuation_date',ascending=False).copy()
                if selected_history.empty:st.info('Esta inversión todavía no tiene valuaciones registradas.')
                else:
                    history_edit=selected_history[['id','valuation_date','value','notes']].rename(columns={'id':'ID','valuation_date':'Fecha','value':'Valor','notes':'Nota'});history_edit['Fecha']=history_edit['Fecha'].dt.date;history_edit['Nota']=history_edit['Nota'].fillna('');history_edit['Eliminar']=False
                    st.caption('Haz doble clic en una celda para modificarla. Después pulsa “Guardar cambios”.')
                    edited_history=st.data_editor(history_edit,hide_index=True,use_container_width=True,num_rows='fixed',disabled=['ID'],column_config={'ID':st.column_config.NumberColumn('ID',format='%d'),'Fecha':st.column_config.DateColumn('Fecha',format='YYYY-MM-DD',required=True),'Valor':st.column_config.NumberColumn('Valor',format='$ %.2f',min_value=0.0,required=True),'Nota':st.column_config.TextColumn('Nota'),'Eliminar':st.column_config.CheckboxColumn('Eliminar')},key=f'history_editor_{history_id}')
                    confirm_value_delete=st.checkbox('Confirmo la eliminación de los renglones marcados.',key=f'confirm_history_delete_{history_id}');save_history=st.button('Guardar cambios',type='primary',key=f'save_history_{history_id}')
                    if save_history:
                        rows_to_delete=edited_history[edited_history['Eliminar']==True]
                        if not rows_to_delete.empty and not confirm_value_delete:st.error('Marcaste valuaciones para eliminar. Confirma su eliminación antes de guardar.')
                        else:
                            try:
                                db=client()
                                for _,edited_row in edited_history.iterrows():
                                    record_id=int(edited_row['ID'])
                                    if bool(edited_row['Eliminar']):db.table('investment_valuations').delete().eq('id',record_id).execute()
                                    else:
                                        corrected_date=pd.to_datetime(edited_row['Fecha']).date().isoformat();corrected_value=float(edited_row['Valor']);corrected_note='' if pd.isna(edited_row['Nota']) else str(edited_row['Nota']);db.table('investment_valuations').update({'valuation_date':corrected_date,'value':corrected_value,'notes':corrected_note}).eq('id',record_id).execute()
                                st.success('Histórico actualizado.');st.rerun()
                            except Exception as e:st.error('No se pudieron guardar los cambios. Verifica que no haya dos valuaciones de la misma inversión en la misma fecha. '+str(e))
        if not valuations.empty:
            st.markdown('### Evolución por inversión')
            st.caption('Cada inversión tiene su propia escala. Acerca o aleja con la rueda del mouse y haz doble clic para restablecer la vista.')

            investment_colors=[BLUE,SKY,RED,'#FCA5A5',GOLD,GREEN,'#8B5CF6','#06B6D4']

            for idx,(_,inv) in enumerate(investments.sort_values(['institution','product']).iterrows()):
                inv_history=valuations[valuations['investment_id']==inv['id']].copy().sort_values('valuation_date')
                if inv_history.empty:
                    continue

                # La gráfica muestra una sola valuación por día: la última registrada.
                # El histórico conserva todos los registros disponibles.
                inv_chart=inv_history.copy()
                inv_chart['chart_date']=inv_chart['valuation_date'].dt.normalize()
                inv_chart=inv_chart.sort_values('valuation_date').groupby('chart_date',as_index=False).tail(1).sort_values('chart_date')

                current_value=float(inv_chart.iloc[-1]['value'])
                previous_value=float(inv_chart.iloc[-2]['value']) if len(inv_chart)>1 else pd.NA
                last_change=current_value-previous_value if pd.notna(previous_value) else pd.NA
                last_pct=(current_value/previous_value-1) if pd.notna(previous_value) and previous_value else pd.NA

                st.markdown(f"#### {inv['label']}")
                m1,m2,m3=st.columns(3)
                m1.metric('Valor actual',money(current_value))
                if pd.isna(last_change):
                    m2.metric('Cambio último','—')
                    m3.metric('% último','—')
                else:
                    m2.metric('Cambio último',money(last_change),delta=money(last_change),delta_color='normal')
                    m3.metric('% último',f'{last_pct:.2%}',delta=f'{last_pct:.2%}',delta_color='normal')

                fig=go.Figure()
                fig.add_trace(go.Scatter(
                    x=inv_chart['chart_date'],
                    y=inv_chart['value'],
                    mode='lines+markers',
                    name=inv['label'],
                    line=dict(width=3,color=investment_colors[idx%len(investment_colors)]),
                    marker=dict(size=8,color=investment_colors[idx%len(investment_colors)]),
                    hovertemplate='%{x|%d %b %Y}<br><b>$%{y:,.2f}</b><extra></extra>'
                ))
                fig.update_layout(
                    dragmode='zoom',
                    hovermode='x unified',
                    showlegend=False,
                    height=330,
                    margin=dict(l=10,r=10,t=15,b=10)
                )
                fig.update_xaxes(
                    title='Fecha',
                    tickformat='%d %b %Y',
                    dtick=86400000,
                    fixedrange=False
                )
                fig.update_yaxes(
                    title='Valor',
                    tickprefix='$',
                    tickformat=',.0f',
                    autorange=True,
                    fixedrange=False
                )
                st.plotly_chart(style(fig),use_container_width=True,config=PLOT_CONFIG)

                if idx < len(investments)-1:
                    st.divider()
        with st.expander('Eliminar inversión'):
            delete_options={f"{row['label']} · ID {int(row['id'])}":int(row['id']) for _,row in investments.iterrows()}
            with st.form('delete_investment'):
                selected_delete=st.selectbox('Inversión a eliminar',list(delete_options));confirm_delete=st.checkbox('Entiendo que también se eliminará todo su historial de valuaciones.');delete_sent=st.form_submit_button('Eliminar definitivamente')
            if delete_sent:
                if not confirm_delete:st.error('Marca la casilla de confirmación antes de eliminar.')
                else:
                    try:client().table('investments').delete().eq('id',delete_options[selected_delete]).execute();st.success('Inversión e historial eliminados.');st.rerun()
                    except Exception as e:st.error('No se pudo eliminar la inversión. '+str(e))
    with st.expander('Agregar inversión',expanded=investments.empty):
        with st.form('investment'):
            c1,c2=st.columns(2);institution=c1.text_input('Institución');product=c2.text_input('Producto');owner=c1.text_input('Titular');asset=c2.selectbox('Tipo de activo',['Efectivo','Deuda','Renta variable','Fondo','Inmueble','Otro']);balance=c1.number_input('Valor inicial',min_value=0.0);rate=c2.number_input('Tasa anual esperada %',min_value=0.0);opened=st.date_input('Fecha inicial');sent=st.form_submit_button('Guardar',type='primary')
        if sent:
            try:
                saved=insert_one('investments',{'institution':institution,'product':product,'owner':owner,'asset_type':asset,'balance':balance,'annual_rate':rate,'opened_on':opened.isoformat()});inv_id=saved[0]['id'] if isinstance(saved,list) else saved['id'];insert_one('investment_valuations',{'investment_id':inv_id,'valuation_date':opened.isoformat(),'value':balance,'notes':'Valuación inicial'});st.success('Inversión guardada');st.rerun()
            except Exception as e:st.error(str(e))
else:
    st.title('Importar movimientos de Alzex');st.write('Carga un CSV completo o mensual. La huella de cada movimiento evita duplicados.');uploaded=st.file_uploader('Archivo CSV',type=['csv'])
    if uploaded:
        try:
            parsed=parse_alzex(uploaded.getvalue(),uploaded.name);existing={r['fingerprint'] for r in fetch('movements')};new=parsed[~parsed.fingerprint.isin(existing)];c1,c2,c3=st.columns(3);c1.metric('Movimientos válidos',f'{len(parsed):,}');c2.metric('Nuevos',f'{len(new):,}');c3.metric('Duplicados',f'{len(parsed)-len(new):,}');st.dataframe(new.head(50),use_container_width=True,hide_index=True)
            if st.button('Confirmar importación',type='primary',disabled=new.empty):insert_rows('movements',new.where(pd.notna(new),None).to_dict('records'));insert_one('imports',{'file_name':uploaded.name,'row_count':len(parsed),'new_rows':len(new),'duplicate_rows':len(parsed)-len(new)});st.success('Importación completada');st.cache_data.clear();st.rerun()
        except Exception as e:st.error(f'No se pudo leer el archivo: {e}')
