"""Reporte PDF de secciones y modelo 3D: usa las mismas secciones del ZDA."""
from io import BytesIO
from pathlib import Path
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
from reportlab.lib.utils import ImageReader

NAVY = '#1C3552'

def _section_image(sec, key, nominal, color, xlim, zlim):
    entry = sec[key]
    fig, ax = plt.subplots(figsize=(4.3, 3.45), dpi=145)
    fig.patch.set_facecolor('white')
    for triangle in entry.get('triangles', []):
        a = np.asarray(triangle, dtype=float)
        if len(a) >= 3:
            ax.fill(a[:, 0], a[:, 1], color=color, alpha=.045)
            ax.plot(np.r_[a[:, 0], a[0, 0]], np.r_[a[:, 1], a[0, 1]], color=color, alpha=.35, lw=.48)
    boundary = np.asarray(entry.get('boundary', []), dtype=float)
    if boundary.ndim == 2 and len(boundary) >= 3:
        ax.fill(boundary[:, 0], boundary[:, 1], color=color, alpha=.10)
        ax.plot(boundary[:, 0], boundary[:, 1], color=color, lw=2)
    nom = np.asarray(nominal, dtype=float)
    if nom.ndim == 2 and len(nom) >= 3:
        ax.plot(nom[:, 0], nom[:, 1], color='#64748B', lw=1, ls='--')
    points = np.asarray(entry.get('points', []), dtype=float)
    if points.ndim == 2 and len(points):
        ax.scatter(points[:, 0], points[:, 1], color=color, s=7, zorder=4)
    ax.set_xlim(*xlim); ax.set_ylim(*zlim)
    ax.set_aspect('equal', adjustable='box')
    ax.set_xlabel('X (m)', fontsize=8); ax.set_ylabel('Z (m)', fontsize=8)
    ax.tick_params(labelsize=7); ax.grid(alpha=.16)
    fig.tight_layout(pad=.8)
    out = BytesIO(); fig.savefig(out, format='png', dpi=145, facecolor='white')
    plt.close(fig); out.seek(0)
    return out

def _text(c, x, y, value, size=9, bold=False, color=NAVY):
    c.setFillColor(HexColor(color)); c.setFont('Helvetica-Bold' if bold else 'Helvetica', size)
    c.drawString(x, y, str(value))

def _table(c, x, top, width, title, headings, rows, widths=None, bold_last=False):
    c.setFillColor(HexColor('#E9F1F8')); c.roundRect(x, top-23, width, 23, 3, fill=1, stroke=0)
    _text(c, x+9, top-16, title, 10, True)
    n = len(headings); widths = widths or [width/n]*n
    y = top-25
    c.setFillColor(HexColor('#F2F6FA')); c.rect(x, y-25, width, 25, fill=1, stroke=0)
    pos=x
    for label, w in zip(headings,widths):
        _text(c, pos+5, y-17, label, 10.5, True); pos+=w
    y-=25
    for row_idx, row in enumerate(rows):
        c.setStrokeColor(HexColor('#D6E0E9')); c.line(x,y-26,x+width,y-26)
        pos=x
        for value,w in zip(row,widths):
            _text(c,pos+5,y-18,value,11,bold_last and row_idx==len(rows)-1); pos+=w
        y-=27
    return y

def generar_reporte_pdf(mask_volume, masks_result, fig_3d, ciclo, logo_path=None, fecha=None, equipo=None, operador=None):
    """PDF horizontal A3; no reprocesa el ZDA ni modifica los indicadores."""
    if not mask_volume.get('complete'):
        raise ValueError('El ciclo no tiene integración volumétrica completa para generar el reporte.')
    sections = mask_volume['sections']; first,last=sections[0],sections[-1]
    if fig_3d is None:
        raise ValueError('El modelo 3D no está disponible para este ciclo.')
    # Exportamos exactamente la figura 3D del dashboard, no una imagen de ejemplo.
    try:
        fig_png = fig_3d.to_image(format='png', width=1300, height=880, scale=1.5)
    except Exception as exc:
        raise RuntimeError('No se pudo exportar la vista 3D. Instala kaleido (requirements.txt) y comprueba que Chrome/Chromium esté disponible para Kaleido 1.x.') from exc
    from reportlab.lib.pagesizes import A3, landscape
    W,H=landscape(A3); buf=BytesIO(); c=canvas.Canvas(buf,pagesize=(W,H))
    margin=20; gap=8; usable=W-2*margin; panel=(usable-3*gap)/4
    c.setFillColor(HexColor(NAVY)); c.rect(0,H-68,W,68,fill=1,stroke=0)
    _text(c,margin+10,H-28,f'Ciclo {ciclo}  |  Eficiencia de Perforación',17,True,'#FFFFFF')
    # Datos del ciclo seleccionado, recibidos desde el ZDA; nunca son valores fijos.
    def dato_cabecera(valor):
        texto = str(valor).strip() if valor is not None else ''
        return texto if texto and texto.lower() not in ('nan', 'none', 'nat') else 'SIN DATO'
    _text(c,margin+10,H-49,
          f'Fecha: {dato_cabecera(fecha)}  |  Equipo: {dato_cabecera(equipo)}  |  Operador: {dato_cabecera(operador)}',
          11,False,'#FFFFFF')
    logo=Path(logo_path) if logo_path else Path(__file__).with_name('logo_reporte.png')
    if logo.is_file():
        c.drawImage(str(logo),W-86,H-59,width=50,height=45,preserveAspectRatio=True,mask='auto')
    coords=[]
    for sec in (first,last):
        for key in ('programado','real'):
            entry=sec[key]; coords.extend(entry.get('points',[])); coords.extend(entry.get('boundary',[]))
    coords.extend(masks_result.get('nominal',[]))
    coords=np.asarray([p[:2] for p in coords if len(p)>=2 and np.isfinite(p[0]) and np.isfinite(p[1])],float)
    xlim=(float(coords[:,0].min())-.4,float(coords[:,0].max())+.4) if len(coords) else (-4,4)
    zlim=(float(coords[:,1].min())-.4,float(coords[:,1].max())+.4) if len(coords) else (-1,6)
    top=H-78; card_h=266
    specs=[(first,'programado','#315FCB','Sección programada | Collar'),(first,'real','#00A878','Sección ejecutada | Collar'),(last,'programado','#E33D45','Sección programada | Fondo'),(last,'real','#ED7D20','Sección ejecutada | Fondo')]
    for idx,(sec,key,color,title) in enumerate(specs):
        x=margin+idx*(panel+gap)
        c.setStrokeColor(HexColor('#D7E1EA')); c.roundRect(x,top-card_h,panel,card_h,4,stroke=1,fill=0)
        c.setFillColor(HexColor('#EAF2FA')); c.roundRect(x,top-24,panel,24,4,fill=1,stroke=0)
        _text(c,x+8,top-16,f'{title} ({sec["depth_m"]:.2f} m)',9,True)
        image=_section_image(sec,key,masks_result.get('nominal',[]),color,xlim,zlim)
        c.drawImage(ImageReader(image),x+5,top-card_h+26,width=panel-10,height=card_h-56,preserveAspectRatio=True,anchor='c')
        entry=sec[key]
        _text(c,x+9,top-card_h+11,f'{len(entry.get("points",[]))} puntos | {len(entry.get("triangles",[]))} triángulos',8)
    lower_top=top-card_h-12; left_w=usable*.53; right_x=margin+left_w+10; right_w=usable-left_w-10
    lower_h=lower_top-38
    c.setStrokeColor(HexColor('#D7E1EA')); c.roundRect(margin,lower_top-lower_h,left_w,lower_h,4,stroke=1,fill=0)
    c.setFillColor(HexColor('#EAF2FA')); c.roundRect(margin,lower_top-24,left_w,24,4,fill=1,stroke=0)
    _text(c,margin+10,lower_top-16,'Modelo volumétrico 3D | Programado vs. ejecutado',10,True)
    c.drawImage(ImageReader(BytesIO(fig_png)),margin+5,lower_top-lower_h+9,width=left_w-10,height=lower_h-40,preserveAspectRatio=True,anchor='c')
    def fmt(x): return f'{x:,.2f}' if x is not None else 'N/D'
    area_rows=[]
    for label,sec in [('Collar',first),('Fondo',last)]:
        p=sec['programado'].get('area_m2'); r=sec['real'].get('area_m2')
        area_rows.append([f'{label} ({sec["depth_m"]:.2f} m)',fmt(p),fmt(r),fmt(r-p) if p is not None and r is not None else 'N/D',f'{100*(r-p)/p:+.1f}%' if p and r is not None else 'N/D'])
    widths=[right_w*.29,right_w*.18,right_w*.17,right_w*.18,right_w*.18]
    y=_table(c,right_x,lower_top,right_w,'Área de sección (m²)', ['Sección','Program.','Ejecutado','Diferencia','Variación'],area_rows,widths)-9
    p=float(mask_volume['programado_m3']); r=float(mask_volume['real_m3']); pct=100*(r-p)/p if p>0 else None
    fuera=float(mask_volume["outside_m3"]); no_cubierto=float(mask_volume["not_covered_m3"])
    base=float(mask_volume["programado_m3"]); dgt=fuera+no_cubierto
    def pct_spatial(value):
        return f"{100*value/base:.2f}%" if base>0 else "N/D"
    spatial_rows=[
        ["Fuera del programado",fmt(fuera),pct_spatial(fuera)],
        ["No cubierto",fmt(no_cubierto),pct_spatial(no_cubierto)],
        ["Desviación geométrica total",fmt(dgt),pct_spatial(dgt)],
    ]
    y=_table(c,right_x,y,right_w,'Diferencias espaciales (m³)',
             ['Indicador','Volumen (m³)','% del programado'],spatial_rows,
             [right_w*.49,right_w*.23,right_w*.28],bold_last=True)-9
    y=_table(c,right_x,y,right_w,'Volumen integrado (m³)',['Intervalo','Program.','Ejecutado','Diferencia','Variación'],[[f'0-{last["depth_m"]:.2f} m',fmt(p),fmt(r),f'{r-p:+.2f}',f'{pct:+.2f}%' if pct is not None else 'N/D']],widths)-8
    _text(c,right_x+5,y-62,'Leyenda',10,True)
    for i,(color,label) in enumerate([('#315FCB','Programado | collar'),('#00A878','Ejecutado | collar'),('#E33D45','Programado | fondo'),('#ED7D20','Ejecutado | fondo')]):
        xx=right_x+5+(i%2)*(right_w/2); yy=y-80-(i//2)*20
        c.setFillColor(HexColor(color)); c.rect(xx,yy-3,12,9,fill=1,stroke=0)
        _text(c,xx+17,yy-2,label,8)
    _text(c,right_x+5,y-134,'Visualización exploratoria: geometría inferida de perforación.',8)
    _text(c,right_x+5,y-148,'No corresponde a sobrerotura medida después de la voladura.',8)
    c.setStrokeColor(HexColor(NAVY)); c.line(margin,29,W-margin,29)
    _text(c,margin,15,'BNV - El Brocal  |  Análisis de perforación con datos ZDA',8)
    c.setFont('Helvetica',8); c.drawRightString(W-margin,15,'Herramienta del sistema de gestión operativa')
    c.showPage(); c.save(); buf.seek(0); return buf.getvalue()
