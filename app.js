/* Salida VIN 2.0 · app.js
   Fases configurables por lista, columnas libres desde Excel, modo rápido, escáner,
   notas/fotos e historial. Datos en Supabase (REST). Funciona sin señal: los cambios
   se guardan en el celular y se suben cuando vuelve la conexión. */
(function () {
'use strict';
const VERSION = '2.1.1';
const CFG = window.CFG || {};
const URL_ = (CFG.SUPABASE_URL || '').replace(/\/+$/, '');
const KEY = CFG.SUPABASE_KEY || '';
const CONFIGURED = /^https:\/\//.test(URL_) && KEY.length > 20;
const $ = id => document.getElementById(id);

/* ======================= idioma ======================= */
const EN = {
  'Salida VIN':'VIN Checkout','Lista':'List','carros':'cars','terminados':'done','sin empezar':'not started',
  'Últimos 4 dígitos del VIN':'Last 4 digits of the VIN','⚡ Modo rápido':'⚡ Quick mode','Apagado':'Off',
  'Escribe 4 dígitos (5 si hay dos iguales). Toca una fase para marcarla o desmarcarla; ⋯ para nota o foto. En modo rápido, Enter o el escáner marcan la fase elegida.':'Type 4 digits (5 if two match). Tap a stage to mark or unmark it; ⋯ for a note or photo. In quick mode, Enter or the scanner marks the chosen stage.',
  'Ver más':'Show more','Lista y respaldo':'List & backup','Exportar a Excel':'Export to Excel','Copiar resumen':'Copy summary',
  '🔒 Admin':'🔒 Admin','🔓 Salir de admin':'🔓 Exit admin','Subir Excel':'Upload Excel','Fases de esta lista':'Stages of this list','Columnas':'Columns',
  'Reiniciar lista':'Reset list','Archivar lista':'Archive list','Desarchivar lista':'Unarchive list','Cancelar':'Cancel','Sí':'Yes','Guardar':'Save','Cerrar':'Close',
  '¿Cómo te llamas?':"What's your name?",'Se guarda en este celular y sale junto a lo que marques.':'Saved on this phone and shown next to what you mark.',
  'PIN de administrador':'Admin PIN','Entrar':'Enter','Columnas del Excel':'Excel columns','Nombre de la lista':'List name','Tipo (plantilla de fases)':'Type (stage template)',
  'Agregar a la lista actual':'Add to current list','Crear lista nueva':'Create new list','Se pueden marcar en cualquier orden. Quitar una fase la oculta; lo ya marcado no se borra.':'Stages can be marked in any order. Removing a stage hides it; what was marked is kept.',
  '+ Agregar fase':'+ Add stage','Guardar también como plantilla de este tipo':'Also save as the template for this type','Columnas de esta lista':'Columns of this list',
  'Nota':'Note','📷 Agregar foto':'📷 Add photo','Historial':'History','Guardar nota':'Save note','Escanear VIN':'Scan VIN',
  'Falta configurar Supabase: abre config.js y pon SUPABASE_URL y SUPABASE_KEY.':'Supabase is not set up: open config.js and fill in SUPABASE_URL and SUPABASE_KEY.',
  '⛶ Pantalla completa':'⛶ Full screen','✕ Salir':'✕ Exit','Conectando…':'Connecting…','En vivo · compartido':'Live · shared','Sin conexión':'Offline',
  'cambios por subir':'changes to upload','Sin configurar':'Not set up','Error del servidor':'Server error',
  'Escaneado':'Scanned','Llevado a carga':'Taken to charging','Washing':'Washing','Picking list':'Picking list','Inspeccionado':'Inspected',
  'Destino':'Destination','Modelo':'Model','Color':'Colour','Ubicación':'Location','Despacho':'Dispatch','Quick in / Quick out':'Quick in / Quick out','Otra':'Other',
  'Archivadas':'Archived','(archivada)':'(archived)','No hay ningún VIN que termine en':'No VIN ends in','en esta lista.':'in this list.','Hay':'There are',
  'carros que terminan en':'cars ending in','Confirma con el VIN completo o escribe un dígito más.':'Check the full VIN or type one more digit.',
  'Toca para marcar':'Tap to mark','Todas las fases':'All stages','Sin terminar':'Not finished','Terminados':'Finished','Sin empezar':'Not started','Todos':'All',
  'Les falta':'Missing','La tienen':'Done','Filtrar por…':'Filter by…','Todos los valores':'All values','No hay carros aquí.':'No cars here.',
  'Modo rápido':'Quick mode','cada VIN que confirmes (Enter o escáner) se marca en':'every VIN you confirm (Enter or scanner) is marked as',
  'Marcar':'Mark','ya estaba marcado':'was already marked','¿Desmarcar':'Unmark','en':'on','Desmarcado':'Unmarked',
  'Preparando Excel…':'Preparing Excel…','No se pudo cargar el lector de Excel; revisa la conexión':"Couldn't load the Excel tool; check the connection",
  'Excel listo':'Excel ready','No se pudo descargar el Excel':"Couldn't download the Excel",'Carros':'Cars','Resumen':'Summary','Exportado':'Exported','Total':'Total',
  'fecha':'date','hora':'time','por':'by','Fase':'Stage','Acción':'Action','Fecha':'Date','Hora':'Time','Persona':'Person','marcó':'marked','desmarcó':'unmarked','nota':'note','foto':'photo',
  'Terminados (todas las fases)':'Finished (all stages)','copiado':'copied','No se pudo copiar':"Couldn't copy",
  'No encontré ningún VIN en ese archivo':"I couldn't find any VIN in that file",'No se pudo leer el archivo':"Couldn't read the file",
  'vehículos en la hoja':'vehicles in sheet','VIN en la columna':'VIN in column','no están en la lista actual':'not in the current list',
  'Ver en tarjeta':'Show on card','Solo en detalle':'Detail only','No guardar':"Don't save",'Columna':'Column',
  'Lista creada':'List created','vehículos agregados':'vehicles added','No se pudo guardar; revisa la conexión':"Couldn't save; check the connection",
  'Lista nueva':'New list','¿Desmarcar TODAS las fases de todos los carros de':'Unmark ALL stages of every car in','para todos? Los carros y el historial se mantienen.':'for everyone? Cars and history are kept.',
  'Lista reiniciada':'List reset','¿Archivar':'Archive','? Sigue guardada y puedes consultarla en "Archivadas".':'? It stays saved under "Archived".',
  'Lista archivada':'List archived','Lista activa otra vez':'List active again','PIN incorrecto':'Wrong PIN','Modo admin activado':'Admin mode on','Cambiar nombre':'Change name',
  'Este navegador no permite pantalla completa aquí':"This browser doesn't allow full screen here",'Hay una versión nueva; se actualizó':'A new version was installed',
  'Borrar':'Clear','Escanear código':'Scan code','Clave de Supabase inválida (revisa config.js)':'Invalid Supabase key (check config.js)',
  'Fases guardadas':'Stages saved','Columnas guardadas':'Columns saved','Fase nueva':'New stage','Subir':'Up','Bajar':'Down','Quitar':'Remove',
  'Nota guardada':'Note saved','Subiendo foto…':'Uploading photo…','Foto guardada':'Photo saved','No se pudo subir la foto':"Couldn't upload the photo",
  'Las fotos necesitan conexión':'Photos need a connection','El historial necesita conexión':'History needs a connection','Sin cambios todavía.':'No changes yet.',
  'Hecho':'Done','Pendiente':'Pending','Abriendo cámara…':'Opening camera…','Apunta al código de barras del VIN':'Point at the VIN barcode',
  'No se pudo abrir la cámara. Revisa el permiso en Ajustes > Safari > Cámara.':"Couldn't open the camera. Check Settings > Safari > Camera.",
  'No está en esta lista':'Not in this list','Ver detalle':'Open','Casilla':'Checkbox','Código leído':'Code read','Apunta al código QR o de barras del VIN':'Point at the VIN QR code or barcode','Fase':'Stage'
,"Ver carros":"Show cars","Por fase":"By stage","Toca una fase para filtrar":"Tap a stage to filter","de":"of","Ya la tienen":"Already done","Les falta":"Still missing","Quitar":"Clear","Idioma":"Language","Mostrando {n} carros sin terminar":"Showing {n} unfinished cars","Mostrando {n} carros terminados":"Showing {n} finished cars","Mostrando los {n} carros":"Showing all {n} cars","{n} carros ya tienen {s}":"{n} cars already have {s}","A {n} carros les falta {s}":"{n} cars still need {s}"
};
const KO = {
"Salida VIN": "VIN 출고",
"Lista": "목록",
"carros": "대",
"terminados": "완료",
"sin empezar": "시작 전",
"Últimos 4 dígitos del VIN": "VIN 마지막 4자리",
"⚡ Modo rápido": "⚡ 빠른 모드",
"Apagado": "끔",
"Escribe 4 dígitos (5 si hay dos iguales). Toca una fase para marcarla o desmarcarla; ⋯ para nota o foto. En modo rápido, Enter o el escáner marcan la fase elegida.": "4자리를 입력하세요(같은 번호가 두 개면 5자리). 단계를 눌러 표시하거나 해제하고, ⋯ 를 눌러 메모나 사진을 추가하세요. 빠른 모드에서는 Enter 또는 스캐너로 선택한 단계가 표시됩니다.",
"Ver más": "더 보기",
"Lista y respaldo": "목록 및 백업",
"Exportar a Excel": "Excel로 내보내기",
"Copiar resumen": "요약 복사",
"🔒 Admin": "🔒 관리자",
"🔓 Salir de admin": "🔓 관리자 종료",
"Subir Excel": "Excel 업로드",
"Fases de esta lista": "이 목록의 단계",
"Columnas": "열",
"Reiniciar lista": "목록 초기화",
"Archivar lista": "목록 보관",
"Desarchivar lista": "보관 해제",
"Cancelar": "취소",
"Sí": "예",
"Guardar": "저장",
"Cerrar": "닫기",
"¿Cómo te llamas?": "이름이 무엇인가요?",
"Se guarda en este celular y sale junto a lo que marques.": "이 휴대폰에 저장되며 표시한 항목 옆에 나타납니다.",
"PIN de administrador": "관리자 PIN",
"Entrar": "들어가기",
"Columnas del Excel": "Excel 열",
"Nombre de la lista": "목록 이름",
"Tipo (plantilla de fases)": "유형(단계 템플릿)",
"Agregar a la lista actual": "현재 목록에 추가",
"Crear lista nueva": "새 목록 만들기",
"Se pueden marcar en cualquier orden. Quitar una fase la oculta; lo ya marcado no se borra.": "단계는 순서 없이 표시할 수 있습니다. 단계를 삭제하면 숨겨지며, 이미 표시한 내용은 유지됩니다.",
"+ Agregar fase": "+ 단계 추가",
"Guardar también como plantilla de este tipo": "이 유형의 템플릿으로도 저장",
"Columnas de esta lista": "이 목록의 열",
"Nota": "메모",
"📷 Agregar foto": "📷 사진 추가",
"Historial": "기록",
"Guardar nota": "메모 저장",
"Escanear VIN": "VIN 스캔",
"Falta configurar Supabase: abre config.js y pon SUPABASE_URL y SUPABASE_KEY.": "Supabase가 설정되지 않았습니다: config.js를 열고 SUPABASE_URL과 SUPABASE_KEY를 입력하세요.",
"⛶ Pantalla completa": "⛶ 전체 화면",
"✕ Salir": "✕ 나가기",
"Conectando…": "연결 중…",
"En vivo · compartido": "실시간 · 공유",
"Sin conexión": "오프라인",
"cambios por subir": "건 업로드 대기",
"Sin configurar": "설정 안 됨",
"Error del servidor": "서버 오류",
"Escaneado": "스캔됨",
"Llevado a carga": "충전 이동",
"Washing": "세차",
"Picking list": "피킹 리스트",
"Inspeccionado": "검사 완료",
"Destino": "목적지",
"Modelo": "모델",
"Color": "색상",
"Ubicación": "위치",
"Despacho": "출고",
"Quick in / Quick out": "Quick in / Quick out",
"Otra": "기타",
"Archivadas": "보관됨",
"(archivada)": "(보관됨)",
"No hay ningún VIN que termine en": "끝자리가",
"en esta lista.": "인 VIN이 이 목록에 없습니다.",
"Hay": "총",
"carros que terminan en": "대의 끝자리가",
"Confirma con el VIN completo o escribe un dígito más.": "전체 VIN을 확인하거나 한 자리를 더 입력하세요.",
"Toca para marcar": "눌러서 표시",
"Todas las fases": "모든 단계",
"Sin terminar": "미완료",
"Terminados": "완료",
"Sin empezar": "시작 전",
"Todos": "전체",
"Les falta": "남음",
"La tienen": "완료",
"Filtrar por…": "필터…",
"Todos los valores": "모든 값",
"No hay carros aquí.": "차량이 없습니다.",
"Modo rápido": "빠른 모드",
"cada VIN que confirmes (Enter o escáner) se marca en": "확인한 모든 VIN(Enter 또는 스캐너)이 다음으로 표시됩니다:",
"Marcar": "표시",
"ya estaba marcado": "이미 표시됨",
"¿Desmarcar": "표시 해제:",
"en": "·",
"Desmarcado": "표시 해제됨",
"Preparando Excel…": "Excel 준비 중…",
"No se pudo cargar el lector de Excel; revisa la conexión": "Excel 도구를 불러오지 못했습니다. 연결을 확인하세요",
"Excel listo": "Excel 준비 완료",
"No se pudo descargar el Excel": "Excel을 다운로드하지 못했습니다",
"Carros": "차량",
"Resumen": "요약",
"Exportado": "내보낸 시간",
"Total": "합계",
"fecha": "날짜",
"hora": "시간",
"por": "담당",
"Fase": "단계",
"Acción": "작업",
"Fecha": "날짜",
"Hora": "시간",
"Persona": "담당자",
"marcó": "표시함",
"desmarcó": "해제함",
"nota": "메모",
"foto": "사진",
"Terminados (todas las fases)": "완료(모든 단계)",
"copiado": "복사됨",
"No se pudo copiar": "복사하지 못했습니다",
"No encontré ningún VIN en ese archivo": "이 파일에서 VIN을 찾지 못했습니다",
"No se pudo leer el archivo": "파일을 읽지 못했습니다",
"vehículos en la hoja": "대 (시트)",
"VIN en la columna": "VIN 열:",
"no están en la lista actual": "대는 현재 목록에 없음",
"Ver en tarjeta": "카드에 표시",
"Solo en detalle": "상세에만",
"No guardar": "저장 안 함",
"Columna": "열",
"Lista creada": "목록이 생성되었습니다",
"vehículos agregados": "대 추가됨",
"No se pudo guardar; revisa la conexión": "저장하지 못했습니다. 연결을 확인하세요",
"Lista nueva": "새 목록",
"¿Desmarcar TODAS las fases de todos los carros de": "다음 목록의 모든 차량에서 모든 단계를 해제할까요:",
"para todos? Los carros y el historial se mantienen.": "(모든 사용자에게 적용) 차량과 기록은 유지됩니다.",
"Lista reiniciada": "목록이 초기화되었습니다",
"¿Archivar": "보관할까요:",
"? Sigue guardada y puedes consultarla en \"Archivadas\".": "? '보관됨'에서 계속 볼 수 있습니다.",
"Lista archivada": "목록이 보관되었습니다",
"Lista activa otra vez": "목록이 다시 활성화되었습니다",
"PIN incorrecto": "PIN이 틀렸습니다",
"Modo admin activado": "관리자 모드 켜짐",
"Cambiar nombre": "이름 변경",
"Este navegador no permite pantalla completa aquí": "이 브라우저에서는 전체 화면을 사용할 수 없습니다",
"Hay una versión nueva; se actualizó": "새 버전이 설치되었습니다",
"Borrar": "지우기",
"Escanear código": "코드 스캔",
"Clave de Supabase inválida (revisa config.js)": "Supabase 키가 올바르지 않습니다(config.js 확인)",
"Fases guardadas": "단계가 저장되었습니다",
"Columnas guardadas": "열이 저장되었습니다",
"Fase nueva": "새 단계",
"Subir": "위로",
"Bajar": "아래로",
"Quitar": "해제",
"Nota guardada": "메모가 저장되었습니다",
"Subiendo foto…": "사진 업로드 중…",
"Foto guardada": "사진이 저장되었습니다",
"No se pudo subir la foto": "사진을 업로드하지 못했습니다",
"Las fotos necesitan conexión": "사진은 인터넷 연결이 필요합니다",
"El historial necesita conexión": "기록은 인터넷 연결이 필요합니다",
"Sin cambios todavía.": "아직 변경 사항이 없습니다.",
"Hecho": "완료",
"Pendiente": "대기",
"Abriendo cámara…": "카메라 여는 중…",
"Apunta al código de barras del VIN": "VIN 바코드를 비추세요",
"No se pudo abrir la cámara. Revisa el permiso en Ajustes > Safari > Cámara.": "카메라를 열 수 없습니다. 설정 > Safari > 카메라를 확인하세요.",
"No está en esta lista": "이 목록에 없음",
"Ver detalle": "열기",
"Casilla": "체크박스",
"Código leído": "읽은 코드",
"Apunta al código QR o de barras del VIN": "VIN QR 코드나 바코드를 비추세요",
"Ver carros": "차량 보기",
"Por fase": "단계별",
"Toca una fase para filtrar": "단계를 눌러 필터링",
"de": "/",
"Ya la tienen": "완료함",
"Idioma": "언어",
"Mostrando {n} carros sin terminar": "미완료 차량 {n}대",
"Mostrando {n} carros terminados": "완료 차량 {n}대",
"Mostrando los {n} carros": "전체 차량 {n}대",
"{n} carros ya tienen {s}": "{s} 완료 차량 {n}대",
"A {n} carros les falta {s}": "{s} 남은 차량 {n}대"
};
const ZH = {
"Salida VIN": "VIN 出库",
"Lista": "列表",
"carros": "辆",
"terminados": "已完成",
"sin empezar": "未开始",
"Últimos 4 dígitos del VIN": "VIN 后 4 位",
"⚡ Modo rápido": "⚡ 快速模式",
"Apagado": "关闭",
"Escribe 4 dígitos (5 si hay dos iguales). Toca una fase para marcarla o desmarcarla; ⋯ para nota o foto. En modo rápido, Enter o el escáner marcan la fase elegida.": "输入 4 位数字（如有重复则输入 5 位）。点击阶段即可标记或取消标记；点 ⋯ 添加备注或照片。在快速模式下，按 Enter 或扫描即可标记所选阶段。",
"Ver más": "显示更多",
"Lista y respaldo": "列表与备份",
"Exportar a Excel": "导出到 Excel",
"Copiar resumen": "复制摘要",
"🔒 Admin": "🔒 管理员",
"🔓 Salir de admin": "🔓 退出管理员",
"Subir Excel": "上传 Excel",
"Fases de esta lista": "此列表的阶段",
"Columnas": "列",
"Reiniciar lista": "重置列表",
"Archivar lista": "归档列表",
"Desarchivar lista": "取消归档",
"Cancelar": "取消",
"Sí": "是",
"Guardar": "保存",
"Cerrar": "关闭",
"¿Cómo te llamas?": "你叫什么名字？",
"Se guarda en este celular y sale junto a lo que marques.": "保存在这部手机上，并显示在你标记的内容旁边。",
"PIN de administrador": "管理员 PIN",
"Entrar": "进入",
"Columnas del Excel": "Excel 列",
"Nombre de la lista": "列表名称",
"Tipo (plantilla de fases)": "类型（阶段模板）",
"Agregar a la lista actual": "添加到当前列表",
"Crear lista nueva": "新建列表",
"Se pueden marcar en cualquier orden. Quitar una fase la oculta; lo ya marcado no se borra.": "阶段可按任意顺序标记。删除阶段只会隐藏它，已标记的内容会保留。",
"+ Agregar fase": "+ 添加阶段",
"Guardar también como plantilla de este tipo": "同时保存为此类型的模板",
"Columnas de esta lista": "此列表的列",
"Nota": "备注",
"📷 Agregar foto": "📷 添加照片",
"Historial": "记录",
"Guardar nota": "保存备注",
"Escanear VIN": "扫描 VIN",
"Falta configurar Supabase: abre config.js y pon SUPABASE_URL y SUPABASE_KEY.": "尚未设置 Supabase：请打开 config.js 并填写 SUPABASE_URL 和 SUPABASE_KEY。",
"⛶ Pantalla completa": "⛶ 全屏",
"✕ Salir": "✕ 退出",
"Conectando…": "连接中…",
"En vivo · compartido": "实时 · 共享",
"Sin conexión": "离线",
"cambios por subir": "条更改待上传",
"Sin configurar": "未设置",
"Error del servidor": "服务器错误",
"Escaneado": "已扫描",
"Llevado a carga": "已送充电",
"Washing": "洗车",
"Picking list": "拣货单",
"Inspeccionado": "已检验",
"Destino": "目的地",
"Modelo": "车型",
"Color": "颜色",
"Ubicación": "位置",
"Despacho": "发运",
"Quick in / Quick out": "Quick in / Quick out",
"Otra": "其他",
"Archivadas": "已归档",
"(archivada)": "（已归档）",
"No hay ningún VIN que termine en": "此列表中没有以",
"en esta lista.": "结尾的 VIN。",
"Hay": "共有",
"carros que terminan en": "辆车的尾号为",
"Confirma con el VIN completo o escribe un dígito más.": "请核对完整 VIN 或多输入一位数字。",
"Toca para marcar": "点击标记",
"Todas las fases": "全部阶段",
"Sin terminar": "未完成",
"Terminados": "已完成",
"Sin empezar": "未开始",
"Todos": "全部",
"Les falta": "未做",
"La tienen": "已完成",
"Filtrar por…": "筛选…",
"Todos los valores": "全部值",
"No hay carros aquí.": "这里没有车辆。",
"Modo rápido": "快速模式",
"cada VIN que confirmes (Enter o escáner) se marca en": "每个确认的 VIN（Enter 或扫描）都会标记为",
"Marcar": "标记",
"ya estaba marcado": "已经标记过",
"¿Desmarcar": "取消标记",
"en": "·",
"Desmarcado": "已取消标记",
"Preparando Excel…": "正在准备 Excel…",
"No se pudo cargar el lector de Excel; revisa la conexión": "无法加载 Excel 工具，请检查网络连接",
"Excel listo": "Excel 已就绪",
"No se pudo descargar el Excel": "无法下载 Excel",
"Carros": "车辆",
"Resumen": "摘要",
"Exportado": "导出时间",
"Total": "合计",
"fecha": "日期",
"hora": "时间",
"por": "操作人",
"Fase": "阶段",
"Acción": "操作",
"Fecha": "日期",
"Hora": "时间",
"Persona": "人员",
"marcó": "已标记",
"desmarcó": "已取消",
"nota": "备注",
"foto": "照片",
"Terminados (todas las fases)": "已完成（全部阶段）",
"copiado": "已复制",
"No se pudo copiar": "无法复制",
"No encontré ningún VIN en ese archivo": "在该文件中找不到任何 VIN",
"No se pudo leer el archivo": "无法读取文件",
"vehículos en la hoja": "辆（工作表）",
"VIN en la columna": "VIN 所在列：",
"no están en la lista actual": "辆不在当前列表中",
"Ver en tarjeta": "在卡片上显示",
"Solo en detalle": "仅在详情中",
"No guardar": "不保存",
"Columna": "列",
"Lista creada": "列表已创建",
"vehículos agregados": "辆已添加",
"No se pudo guardar; revisa la conexión": "无法保存，请检查网络连接",
"Lista nueva": "新列表",
"¿Desmarcar TODAS las fases de todos los carros de": "要为所有人取消以下列表中所有车辆的全部阶段吗：",
"para todos? Los carros y el historial se mantienen.": "？车辆和记录会保留。",
"Lista reiniciada": "列表已重置",
"¿Archivar": "归档",
"? Sigue guardada y puedes consultarla en \"Archivadas\".": "？之后仍可在“已归档”中查看。",
"Lista archivada": "列表已归档",
"Lista activa otra vez": "列表已重新启用",
"PIN incorrecto": "PIN 错误",
"Modo admin activado": "管理员模式已开启",
"Cambiar nombre": "更改名字",
"Este navegador no permite pantalla completa aquí": "此浏览器不支持全屏",
"Hay una versión nueva; se actualizó": "已安装新版本",
"Borrar": "清除",
"Escanear código": "扫描条码",
"Clave de Supabase inválida (revisa config.js)": "Supabase 密钥无效（请检查 config.js）",
"Fases guardadas": "阶段已保存",
"Columnas guardadas": "列已保存",
"Fase nueva": "新阶段",
"Subir": "上移",
"Bajar": "下移",
"Quitar": "清除",
"Nota guardada": "备注已保存",
"Subiendo foto…": "正在上传照片…",
"Foto guardada": "照片已保存",
"No se pudo subir la foto": "无法上传照片",
"Las fotos necesitan conexión": "上传照片需要网络连接",
"El historial necesita conexión": "查看记录需要网络连接",
"Sin cambios todavía.": "暂无更改。",
"Hecho": "已完成",
"Pendiente": "待完成",
"Abriendo cámara…": "正在打开相机…",
"Apunta al código de barras del VIN": "对准 VIN 条码",
"No se pudo abrir la cámara. Revisa el permiso en Ajustes > Safari > Cámara.": "无法打开相机。请检查 设置 > Safari > 相机。",
"No está en esta lista": "不在此列表中",
"Ver detalle": "打开",
"Casilla": "复选框",
"Código leído": "已读取的条码",
"Apunta al código QR o de barras del VIN": "对准 VIN 二维码或条形码",
"Ver carros": "查看车辆",
"Por fase": "按阶段",
"Toca una fase para filtrar": "点击阶段进行筛选",
"de": "/",
"Ya la tienen": "已做",
"Idioma": "语言",
"Mostrando {n} carros sin terminar": "显示 {n} 辆未完成车辆",
"Mostrando {n} carros terminados": "显示 {n} 辆已完成车辆",
"Mostrando los {n} carros": "显示全部 {n} 辆车",
"{n} carros ya tienen {s}": "已完成 {s} 的车辆：{n} 辆",
"A {n} carros les falta {s}": "尚未完成 {s} 的车辆：{n} 辆"
};
const DICT = { en: EN, ko: KO, zh: ZH };
let lang = 'es';
try { const nl = (navigator.language || '').toLowerCase().slice(0, 2); lang = localStorage.getItem('lang') || (['es', 'ko', 'zh'].includes(nl) ? nl : 'en'); } catch (e) {}
if (!['es', 'en', 'ko', 'zh'].includes(lang)) lang = 'es';
const tr = s => { if (lang === 'es') return s; const D = DICT[lang]; return (D && D[s] !== undefined) ? D[s] : (EN[s] !== undefined ? EN[s] : s); };
const LOC = () => ({ en: 'en-AU', ko: 'ko-KR', zh: 'zh-CN' })[lang] || 'es';

/* ======================= utilidades ======================= */
const LS = (() => { try { const t = '__t'; localStorage.setItem(t, t); localStorage.removeItem(t); return localStorage; } catch (e) { return null; } })();
const lget = (k, d) => { try { const v = LS && LS.getItem(k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } };
const lset = (k, v) => { try { LS && LS.setItem(k, JSON.stringify(v)); } catch (e) {} };
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[c]));
const ms = v => v ? Date.parse(v) : 0;
const iso = t => t ? new Date(t).toISOString() : null;
const slug = s => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '').slice(0, 30) || 'col';
const okColor = c => /^#[0-9a-fA-F]{6}$/.test(c || '') ? c : '#5d645e';
function hm(t) { const d = new Date(t); return d.toLocaleDateString(LOC(), { day:'2-digit', month:'short' }) + ' ' + d.toLocaleTimeString(LOC(), { hour:'2-digit', minute:'2-digit' }); }
function hh(t) { return new Date(t).toLocaleTimeString(LOC(), { hour:'2-digit', minute:'2-digit' }); }
function toast(m) { const t = $('toast'); t.textContent = m; t.classList.add('on'); clearTimeout(toast.h); toast.h = setTimeout(() => t.classList.remove('on'), 2400); }
function ask(m, fn) { $('dlgT').textContent = m; open_('dlg'); $('dlgY').onclick = () => { close_('dlg'); fn(); }; $('dlgN').onclick = () => close_('dlg'); }
function open_(id) { $(id).classList.add('on'); }
function close_(id) { $(id).classList.remove('on'); }
function vinH(v, n) { n = Math.max(4, n || 4); return esc(v.slice(0, -n)) + '<em>' + esc(v.slice(-n)) + '</em>'; }
function loadScript(src) { return new Promise((res, rej) => { const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = () => rej(new Error('load ' + src)); document.head.appendChild(s); }); }
const VIN_RE = /^[A-HJ-NPR-Z0-9]{17}$/;

/* ======================= Supabase REST ======================= */
async function api(path, opt) {
  opt = opt || {};
  const ctl = new AbortController(); const to = setTimeout(() => ctl.abort(), opt.timeout || 15000);
  try {
    const r = await fetch(URL_ + (opt.raw ? '' : '/rest/v1/') + path, {
      method: opt.method || 'GET', signal: ctl.signal, cache: 'no-store',
      headers: Object.assign({ apikey: KEY, Authorization: 'Bearer ' + KEY }, opt.blob ? { 'Content-Type': opt.blob.type } : { 'Content-Type': 'application/json' }, opt.prefer ? { Prefer: opt.prefer } : {}),
      body: opt.blob ? opt.blob : (opt.body !== undefined ? JSON.stringify(opt.body) : undefined)
    });
    if (!r.ok) { const e = new Error('http ' + r.status); e.status = r.status; throw e; }
    if (r.status === 204) return null;
    const txt = await r.text(); return txt ? JSON.parse(txt) : null;
  } finally { clearTimeout(to); }
}
const enc = encodeURIComponent;

/* ======================= estado ======================= */
let LISTS = lget('c:lists', []);       // [{id,name,type,archived,columns,stages,created_at,updated_at}]
const DEF_TPL = { 'Despacho': [{ id: 'escaneado', label: 'Escaneado', color: '#334155' }, { id: 'carga', label: 'Llevado a carga', color: '#1d5fd6' }, { id: 'washing', label: 'Washing', color: '#0e7490' }, { id: 'picking', label: 'Picking list', color: '#6d28d9' }, { id: 'inspeccionado', label: 'Inspeccionado', color: '#17803f' }], 'Quick in / Quick out': [{ id: 'picking', label: 'Picking list', color: '#6d28d9' }, { id: 'carga', label: 'Llevado a carga', color: '#1d5fd6' }], 'Otra': [{ id: 'escaneado', label: 'Escaneado', color: '#334155' }, { id: 'carga', label: 'Llevado a carga', color: '#1d5fd6' }, { id: 'washing', label: 'Washing', color: '#0e7490' }, { id: 'picking', label: 'Picking list', color: '#6d28d9' }, { id: 'inspeccionado', label: 'Inspeccionado', color: '#17803f' }] };
let TPL = lget('c:tpl', null) || DEF_TPL; // {tipo: [stages]}
let curId = lget('cur', null);
let VEH = [];                           // [{vin, ...campos}]
let MARKS = {};                         // vin -> stage -> {t,by,note,photo}
let OUTBOX = lget('outbox2', []);       // [{list_id,vin,stage,f}]
let ME = lget('me', '');
let ADMIN = !CFG.ADMIN_PIN || lget('admin', false) === true;
let quick = '';                         // id de fase en modo rápido
let synced = false, online = true, serverErr = '', lastSync = null, vehUpdated = null;
let fstage = '', fmode = 'unfinished', fcol = '', fval = '', limit = 60;

const L = () => LISTS.find(l => l.id === curId) || null;
const STAGES = () => { const l = L(); return (l && Array.isArray(l.stages) ? l.stages : []).filter(s => s && s.id && !s.hidden); };
const COLS = () => { const l = L(); return (l && Array.isArray(l.columns) ? l.columns : []).filter(c => c && c.key); };
const done = (v, s) => !!(MARKS[v] && MARKS[v][s] && MARKS[v][s].t);
const nDone = v => STAGES().reduce((n, s) => n + (done(v, s.id) ? 1 : 0), 0);
const finished = v => { const S = STAGES(); return S.length > 0 && S.every(s => done(v, s.id)); };
const pendingKey = new Set();
function rebuildPending() { pendingKey.clear(); for (const x of OUTBOX) pendingKey.add(x.list_id + '|' + x.vin + '|' + x.stage); }
rebuildPending();
function loadCache(id) { VEH = lget('c:veh:' + id, []); MARKS = lget('c:marks:' + id, {}); lastSync = lget('c:sync:' + id, null); vehUpdated = lget('c:vehu:' + id, null); }
function saveCache() { if (!curId) return; lset('c:veh:' + curId, VEH); lset('c:marks:' + curId, MARKS); lset('c:sync:' + curId, lastSync); lset('c:vehu:' + curId, vehUpdated); }

/* ======================= conexión ======================= */
function syncUI() {
  const d = $('dot'), s = $('sync'), n = OUTBOX.length, tail = n ? ' · ' + n + ' ' + tr('cambios por subir') : '';
  if (!CONFIGURED) { d.className = 'dot err'; s.textContent = tr('Sin configurar'); return; }
  if (serverErr) { d.className = 'dot err'; s.textContent = tr('Error del servidor') + tail; return; }
  if (!online || n) { d.className = 'dot off'; s.textContent = (online ? tr('Conectando…') : tr('Sin conexión')) + tail; return; }
  d.className = 'dot'; s.textContent = synced ? tr('En vivo · compartido') : tr('Conectando…');
}

/* ======================= marcar ======================= */
function queue(v, s, f) {
  const prev = OUTBOX.find(x => x.list_id === curId && x.vin === v && x.stage === s);
  if (prev) Object.assign(prev.f, f); else OUTBOX.push({ list_id: curId, vin: v, stage: s, f });
  lset('outbox2', OUTBOX); rebuildPending(); saveCache(); render(); flush();
}
function setMark(v, s, on) {
  if (!curId) return;
  const now = Date.now();
  MARKS[v] = MARKS[v] || {};
  const m = Object.assign({}, MARKS[v][s] || {});
  if (on) { m.t = now; m.by = ME; } else { delete m.t; delete m.by; }
  MARKS[v][s] = m;
  queue(v, s, on ? { done_at: iso(now), done_by: ME || null } : { done_at: null, done_by: ME || null });
}
function setNote(v, s, note) {
  MARKS[v] = MARKS[v] || {}; MARKS[v][s] = Object.assign({}, MARKS[v][s] || {}, { note });
  queue(v, s, { note: note || null, done_by: (MARKS[v][s].by || ME || null) });
}
function stageLabel(id) { const s = (L() && L().stages || []).find(x => x.id === id); return s ? tr(s.label) : id; }
function toggle(v, s, fromList) {
  if (done(v, s)) {
    ask(tr('¿Desmarcar') + ' ' + stageLabel(s) + ' ' + tr('en') + ' ' + v.slice(-4) + '?', () => { setMark(v, s, false); toast(tr('Desmarcado') + ' · ' + stageLabel(s)); });
  } else { setMark(v, s, true); toast('✓ ' + v.slice(-4) + ' · ' + stageLabel(s)); if (!fromList) { $('q').value = ''; search(); $('q').focus(); } }
}

let flushing = false;
async function flush() {
  if (flushing || !CONFIGURED || !OUTBOX.length) { syncUI(); return; }
  flushing = true;
  try {
    while (OUTBOX.length) {
      const x = OUTBOX[0];
      try {
        await api('marks?on_conflict=list_id,vin,stage', { method: 'POST', prefer: 'resolution=merge-duplicates,return=minimal', body: [Object.assign({ list_id: x.list_id, vin: x.vin, stage: x.stage }, x.f)] });
        OUTBOX.shift(); lset('outbox2', OUTBOX); rebuildPending(); online = true; serverErr = '';
      } catch (e) {
        if (e.status === 401 || e.status === 403) { serverErr = 'auth'; toast(tr('Clave de Supabase inválida (revisa config.js)')); break; }
        if (e.status === 409 || e.status === 400) { OUTBOX.shift(); lset('outbox2', OUTBOX); rebuildPending(); continue; }
        if (e.status) serverErr = String(e.status); else online = false;
        break;
      }
    }
  } finally { flushing = false; syncUI(); }
}

/* ======================= leer del servidor ======================= */
function applyRows(rows) {
  let ch = false;
  for (const r of rows || []) {
    if (r.list_id !== curId) continue;
    if (!lastSync || ms(r.updated_at) > ms(lastSync)) lastSync = r.updated_at;
    if (pendingKey.has(r.list_id + '|' + r.vin + '|' + r.stage)) continue;
    const m = {}; if (r.done_at) { m.t = ms(r.done_at); m.by = r.done_by || ''; } if (r.note) m.note = r.note; if (r.photo) m.photo = r.photo;
    const old = MARKS[r.vin] && MARKS[r.vin][r.stage];
    if (!old || old.t !== m.t || old.by !== m.by || old.note !== m.note || old.photo !== m.photo) { MARKS[r.vin] = MARKS[r.vin] || {}; MARKS[r.vin][r.stage] = m; ch = true; }
  }
  return ch;
}
async function pullLists() {
  const rows = await api('lists?select=id,name,type,archived,columns,stages,created_at,updated_at&order=created_at.desc');
  LISTS = rows || []; lset('c:lists', LISTS); synced = true;
  const t = await api('settings?key=eq.templates&select=value'); if (t && t[0] && t[0].value) { TPL = t[0].value; lset('c:tpl', TPL); }
  if (!L()) { const a = LISTS.find(l => !l.archived) || LISTS[0]; if (a) await openList(a.id, true); }
  fillLists(); fillQuick(); fillFilters(); render();
}
async function pullVehicles() {
  const l = L(); if (!l) return;
  if (VEH.length && vehUpdated === l.updated_at) return;
  const rows = await api('lists?id=eq.' + enc(curId) + '&select=vehicles,updated_at');
  if (rows && rows[0]) {
    VEH = (Array.isArray(rows[0].vehicles) ? rows[0].vehicles : []).filter(r => r && typeof r.vin === 'string');
    vehUpdated = rows[0].updated_at; saveCache(); fillFilters(); render();
  }
}
async function pullMarks(full) {
  if (!curId) return; const id = curId;
  let q = 'marks?list_id=eq.' + enc(id) + '&select=list_id,vin,stage,done_at,done_by,note,photo,updated_at&order=updated_at.asc&limit=20000';
  if (!full && lastSync) q += '&updated_at=gt.' + enc(new Date(ms(lastSync) - 5000).toISOString());
  const rows = await api(q); if (id !== curId) return;
  if (full) { const keep = {}; for (const v in MARKS) for (const s in MARKS[v]) if (pendingKey.has(id + '|' + v + '|' + s)) { keep[v] = keep[v] || {}; keep[v][s] = MARKS[v][s]; } MARKS = keep; lastSync = null; }
  const ch = applyRows(rows); synced = true; saveCache(); if (ch || full) render();
}
let tick = 0, polling = false;
async function poll() {
  if (polling || !CONFIGURED || document.hidden) return;
  polling = true;
  try {
    await flush();
    if (tick % 12 === 0) await pullLists();          // listas y fases: cada ~1 min
    await pullVehicles();
    await pullMarks(tick % 720 === 0);               // recarga completa: cada ~1 h
    online = true; if (serverErr !== 'auth') serverErr = ''; tick++;
  } catch (e) { if (e.status === 401 || e.status === 403) serverErr = 'auth'; else if (e.status) serverErr = String(e.status); else online = false; }
  finally { polling = false; syncUI(); }
}
async function openList(id, quiet) {
  if (id === curId && VEH.length) { fillLists(); return; }
  curId = id; lset('cur', id); loadCache(id); fstage = ''; fmode = 'unfinished'; fcol = ''; fval = ''; limit = 60;
  if (quick && !STAGES().some(s => s.id === quick)) quick = '';
  fillLists(); fillQuick(); fillFilters(); render(); syncUI();
  if (!quiet) { tick = 1; try { await pullVehicles(); await pullMarks(true); online = true; } catch (e) { if (!e.status) online = false; } syncUI(); }
}

/* ======================= pantalla ======================= */
function fillLists() {
  const sel = $('lsel'), act = LISTS.filter(l => !l.archived), arc = LISTS.filter(l => l.archived);
  const opt = l => '<option value="' + esc(l.id) + '">' + esc(l.name) + (l.type ? ' · ' + esc(tr(l.type)) : '') + '</option>';
  sel.innerHTML = act.map(opt).join('') + (arc.length ? '<optgroup label="' + esc(tr('Archivadas')) + '">' + arc.map(opt).join('') + '</optgroup>' : '');
  sel.value = curId || ''; $('lbar').hidden = !LISTS.length;
  const l = L(); $('lname').textContent = l ? l.name + (l.type ? ' · ' + tr(l.type) : '') + (l.archived ? ' ' + tr('(archivada)') : '') : '';
  $('arch').textContent = tr(l && l.archived ? 'Desarchivar lista' : 'Archivar lista');
}
function fillQuick() {
  $('qsel').innerHTML = '<option value="">' + esc(tr('Apagado')) + '</option>' + STAGES().map(s => '<option value="' + esc(s.id) + '">' + esc(tr(s.label)) + '</option>').join('');
  $('qsel').value = quick; const s = STAGES().find(x => x.id === quick), b = $('qban');
  b.hidden = !s; if (s) { b.style.background = okColor(s.color); b.textContent = '⚡ ' + tr('Modo rápido') + ': ' + tr('cada VIN que confirmes (Enter o escáner) se marca en') + ' ' + tr(s.label); }
}
function fillFilters() {
  const S = STAGES(), C = COLS();
  if (fstage && !S.some(s => s.id === fstage)) fstage = '';
  if (fstage ? !['has', 'missing'].includes(fmode) : !['unfinished', 'finished', 'all'].includes(fmode)) fmode = fstage ? 'has' : 'unfinished';
  if (fcol && !colChoices().some(o => o.c.key === fcol && o.vals.includes(fval))) { fcol = ''; fval = ''; }
  void C;
}
const cval = (x, k) => String(x[k] == null ? '' : x[k]).trim();
function colChoices() {
  return COLS().filter(c => c.show !== 'no').map(c => ({ c, vals: [...new Set(VEH.map(x => cval(x, c.key)))].filter(Boolean).sort() }))
    .filter(o => o.vals.length >= 2 && o.vals.length <= 40 && o.vals.length < VEH.length).slice(0, 3);
}
function fgOn(c) { const n = parseInt(okColor(c).slice(1), 16); return (0.299 * (n >> 16) + 0.587 * (n >> 8 & 255) + 0.114 * (n & 255)) > 160 ? '#141714' : '#fff'; }
function abbrs() {
  const S = STAGES(), AB = {}, clean = s => tr(s.label).replace(/[^A-Za-zÁÉÍÓÚÑáéíóúñ0-9]/g, '').toUpperCase() || tr(s.label).slice(0, 2);
  S.forEach(s => { const c = clean(s), one = c.slice(0, 1); AB[s.id] = S.filter(o => clean(o).slice(0, 1) === one).length > 1 ? c.slice(0, 2) : one; });
  return AB;
}
function drawFilters(shown) {
  const S = STAGES(), n = S.length, AB = abbrs(), base = (fcol && fval) ? VEH.filter(x => cval(x, fcol) === fval) : VEH;
  let fin = 0; const cnt = {}; S.forEach(s => cnt[s.id] = 0);
  for (const x of base) { let k = 0; for (const s of S) if (done(x.vin, s.id)) { cnt[s.id]++; k++; } if (n && k === n) fin++; }
  const tot = base.length;
  $('seg').innerHTML = [['unfinished', 'Sin terminar', tot - fin], ['finished', 'Terminados', fin], ['all', 'Todos', tot]].map(m => '<button type="button" data-seg="' + m[0] + '" aria-pressed="' + (!fstage && fmode === m[0]) + '">' + esc(tr(m[1])) + '<b>' + m[2] + '</b></button>').join('');
  $('cards').innerHTML = S.map(s => { const c = okColor(s.color), p = tot ? Math.round(cnt[s.id] / tot * 100) : 0;
    return '<button type="button" class="card" data-card="' + esc(s.id) + '" aria-pressed="' + (fstage === s.id) + '" style="--c:' + c + ';--f:' + fgOn(c) + '"><span class="t"><span class="l" aria-hidden="true">' + esc(AB[s.id]) + '</span>' + esc(tr(s.label)) + '</span><span class="n">' + cnt[s.id] + ' <small>' + esc(tr('de')) + ' ' + tot + '</small></span><span class="bar"><i style="width:' + p + '%"></i></span></button>'; }).join('');
  const st = S.find(s => s.id === fstage), hm = $('hm'); hm.hidden = !st;
  if (st) { const c = okColor(st.color); hm.style.setProperty('--c', c); hm.style.setProperty('--f', fgOn(c));
    hm.innerHTML = '<button type="button" data-hm="has" aria-pressed="' + (fmode === 'has') + '">✓ ' + esc(tr('Ya la tienen')) + ' · ' + cnt[st.id] + '</button><button type="button" data-hm="missing" aria-pressed="' + (fmode === 'missing') + '">' + esc(tr('Les falta')) + ' · ' + (tot - cnt[st.id]) + '</button>'; }
  $('cf').innerHTML = colChoices().map(o => '<div class="cf"><span>' + esc(tr(o.c.label)) + '</span><button type="button" data-cf="' + esc(o.c.key) + '" data-cv="" aria-pressed="' + !(fcol === o.c.key && fval) + '">' + esc(tr('Todos')) + '</button>'
    + o.vals.map(v => '<button type="button" data-cf="' + esc(o.c.key) + '" data-cv="' + esc(v) + '" aria-pressed="' + (fcol === o.c.key && fval === v) + '">' + esc(v) + '</button>').join('') + '</div>').join('');
  let msg = st ? tr(fmode === 'has' ? '{n} carros ya tienen {s}' : 'A {n} carros les falta {s}') : tr(fmode === 'finished' ? 'Mostrando {n} carros terminados' : fmode === 'all' ? 'Mostrando los {n} carros' : 'Mostrando {n} carros sin terminar');
  msg = esc(msg).replace('{n}', '<b>' + shown + '</b>').replace('{s}', '<b>' + esc(st ? tr(st.label) : '') + '</b>');
  const fc = fcol && fval ? COLS().find(c => c.key === fcol) : null;
  if (fc) msg += ' · ' + esc(tr(fc.label)) + ': <b>' + esc(fval) + '</b>';
  $('sum').innerHTML = '<span>' + msg + '</span>' + ((st || fc || fmode !== 'unfinished') ? '<button type="button" data-clr="1">✕ ' + esc(tr('Quitar')) + '</button>' : '');
}
function stats() {
  const S = STAGES(); let fin = 0, none = 0; const cnt = {}; S.forEach(s => cnt[s.id] = 0);
  for (const x of VEH) { let k = 0; for (const s of S) if (done(x.vin, s.id)) { cnt[s.id]++; k++; } if (S.length && k === S.length) fin++; if (!k) none++; }
  $('sTot').textContent = VEH.length; $('sDone').textContent = fin; $('sNone').textContent = none;
}
function fieldTags(x, mode) {
  return COLS().filter(c => mode === 'all' ? c.show !== 'no' : c.show === 'card').map(c => x[c.key] ? '<span class="tag"><b>' + esc(tr(c.label)) + '</b>' + esc(x[c.key]) + '</span>' : '').join(' ');
}
function stageButtons(x) {
  return '<div class="stg">' + STAGES().map(s => {
    const m = MARKS[x.vin] && MARKS[x.vin][s.id] || {}, on = !!m.t, extra = (m.note ? ' 📝' : '') + (m.photo ? ' 📷' : '');
    return '<div class="sbw" style="position:relative"><button class="sb' + (on ? ' on' : '') + '" style="--c:' + okColor(s.color) + '" data-s="' + esc(s.id) + '" data-v="' + esc(x.vin) + '">' + (on ? '✓ ' : '') + esc(tr(s.label))
      + '<small>' + (on ? esc(hh(m.t) + (m.by ? ' · ' + m.by : '')) : esc(tr('Toca para marcar'))) + esc(extra) + '</small></button>'
      + '<button class="sx" style="color:' + (on ? '#fff' : okColor(s.color)) + '" data-d="' + esc(s.id) + '" data-v="' + esc(x.vin) + '" aria-label="' + esc(tr('Nota')) + ' / ' + esc(tr('foto')) + '">⋯</button></div>';
  }).join('') + '</div>';
}
function matches() {
  const q = $('q').value.toUpperCase().replace(/[^A-Z0-9]/g, '');
  if (q.length < 4) return { q, m: [] };
  return { q, m: VEH.filter(x => x.vin.endsWith(q) || (q.length >= 6 && x.vin.includes(q))) };
}
function search() {
  const { q, m } = matches(), r = $('res');
  if (q.length < 4) { r.innerHTML = ''; return m; }
  if (!m.length) { r.innerHTML = '<div class="note">' + esc(tr('No hay ningún VIN que termine en')) + ' ' + esc(q) + ' ' + esc(tr('en esta lista.')) + '</div>'; return m; }
  let h = m.length > 1 ? '<div class="note">' + esc(tr('Hay')) + ' ' + m.length + ' ' + esc(tr('carros que terminan en')) + ' ' + esc(q) + '. ' + esc(tr('Confirma con el VIN completo o escribe un dígito más.')) + '</div>' : '';
  const qs = STAGES().find(s => s.id === quick);
  h += m.slice(0, 6).map(x => '<div class="hit' + (finished(x.vin) ? ' done' : '') + '"><div class="vin">' + vinH(x.vin, Math.min(q.length, 6)) + '</div><div class="fields">' + fieldTags(x, 'all') + '</div>'
    + (qs && m.length === 1 && !done(x.vin, qs.id) ? '<button class="pri" style="background:' + okColor(qs.color) + ';border-color:' + okColor(qs.color) + ';min-height:52px;font-size:17px" data-s="' + esc(qs.id) + '" data-v="' + esc(x.vin) + '">⚡ ' + esc(tr('Marcar')) + ' ' + esc(tr(qs.label)) + '</button>' : '')
    + stageButtons(x) + '</div>').join('');
  r.innerHTML = h; return m;
}
function filtered() {
  return VEH.filter(x => {
    if (fcol && fval && cval(x, fcol) !== fval) return false;
    if (fstage) { const d = done(x.vin, fstage); return fmode === 'all' || (fmode === 'has' ? d : !d); }
    if (fmode === 'all') return true;
    const k = nDone(x.vin), n = STAGES().length;
    return fmode === 'finished' ? (n && k === n) : fmode === 'none' ? k === 0 : !(n && k === n);
  });
}
function render() {
  stats(); search();
  let Lx = filtered();
  if (fstage && fmode === 'has') Lx.sort((a, b) => (MARKS[b.vin][fstage].t || 0) - (MARKS[a.vin][fstage].t || 0));
  drawFilters(Lx.length);
  const S = STAGES(), sh = Lx.slice(0, limit), AB = abbrs();
  $('list').innerHTML = sh.length ? sh.map(x => {
    const sub = COLS().filter(c => c.show === 'card' && x[c.key]).map(c => x[c.key]).join(' · ');
    return '<li class="' + (finished(x.vin) ? 'done' : '') + '"><button class="grow open" data-open="' + esc(x.vin) + '" aria-label="' + esc(tr('Ver detalle')) + ' ' + esc(x.vin) + '"><span class="vin">' + vinH(x.vin) + '</span>' + (sub ? '<span class="sub2">' + esc(sub) + '</span>' : '') + '</button><div class="sq">'
      + S.map(s => { const m = MARKS[x.vin] && MARKS[x.vin][s.id]; const on = !!(m && m.t);
        return '<button class="' + (on ? 'on' : '') + '" style="--c:' + okColor(s.color) + '" data-l="' + esc(s.id) + '" data-v="' + esc(x.vin) + '" aria-label="' + esc(tr(s.label)) + ': ' + esc(tr(on ? 'Hecho' : 'Pendiente')) + '" title="' + esc(tr(s.label)) + (on ? ' ' + esc(hh(m.t)) : '') + '">' + esc(AB[s.id]) + '</button>'; }).join('')
      + '</div></li>'; }).join('') : '<li class="empty">' + esc(tr('No hay carros aquí.')) + '</li>';
  $('more').style.display = Lx.length > limit ? 'block' : 'none'; $('more').textContent = tr('Ver más') + ' (' + (Lx.length - limit) + ')';
}

/* ======================= eventos ======================= */
document.body.addEventListener('click', e => {
  const b = e.target.closest('button'); if (!b) return; const d = b.dataset;
  if (d.s && d.v) toggle(d.v, d.s, false);
  else if (d.l && d.v) toggle(d.v, d.l, true);
  else if (d.d && d.v) openDetail(d.v, d.d);
  else if (d.open) { let n = 4; while (n < d.open.length && VEH.filter(x => x.vin.endsWith(d.open.slice(-n))).length > 1) n++; $('q').value = d.open.slice(-n); search(); $('q').closest('.pad').scrollIntoView({ behavior: 'smooth', block: 'start' }); }
  else if (d.seg) { fstage = ''; fmode = d.seg; limit = 60; render(); }
  else if (d.card) { if (fstage === d.card) { fstage = ''; fmode = 'unfinished'; } else { fstage = d.card; fmode = 'has'; } limit = 60; render(); }
  else if (d.hm) { fmode = d.hm; limit = 60; render(); }
  else if (d.cf !== undefined) { fcol = d.cv ? d.cf : ''; fval = d.cv || ''; limit = 60; render(); }
  else if (d.clr) { fstage = ''; fmode = 'unfinished'; fcol = ''; fval = ''; limit = 60; render(); }
});
function quickMark(x) {
  const s = STAGES().find(z => z.id === quick); if (!s || !x) return false;
  if (done(x.vin, s.id)) toast(x.vin.slice(-4) + ' · ' + tr(s.label) + ' ' + tr('ya estaba marcado'));
  else { setMark(x.vin, s.id, true); toast('⚡ ' + x.vin.slice(-4) + ' · ' + tr(s.label)); }
  return true;
}
$('q').addEventListener('input', () => search());
$('q').addEventListener('keydown', e => { if (e.key === 'Enter') { const m = search(); if (m.length === 1 && quick) { quickMark(m[0]); $('q').value = ''; search(); } } });
$('clr').onclick = () => { $('q').value = ''; search(); $('q').focus(); };
$('qsel').onchange = e => { quick = e.target.value; fillQuick(); search(); };
$('more').onclick = () => { limit += 100; render(); };
$('lsel').onchange = e => { $('q').value = ''; quick = ''; openList(e.target.value); };

/* ---------- detalle de fase: nota, foto, historial ---------- */
let det = null;
function openDetail(v, s) {
  det = { v, s }; const m = MARKS[v] && MARKS[v][s] || {};
  $('sdT').textContent = tr(stageLabel(s)) + ' · ' + v.slice(-6);
  $('sdS').textContent = m.t ? '✓ ' + hm(m.t) + (m.by ? ' · ' + m.by : '') : tr('Pendiente');
  $('sdN').value = m.note || '';
  $('sdP').innerHTML = m.photo && /^https:\/\//.test(m.photo) ? '<a href="' + esc(m.photo) + '" target="_blank" rel="noopener"><img class="ph" src="' + esc(m.photo) + '" alt="' + esc(tr('foto')) + '"></a>' : '';
  open_('sd');
}
$('sdC').onclick = () => close_('sd');
$('sdY').onclick = () => { if (!det) return; setNote(det.v, det.s, $('sdN').value.trim()); close_('sd'); toast(tr('Nota guardada')); };
async function shrink(file) {
  const img = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = URL.createObjectURL(file); });
  const k = Math.min(1, 1600 / Math.max(img.naturalWidth, img.naturalHeight)), c = document.createElement('canvas');
  c.width = Math.round(img.naturalWidth * k); c.height = Math.round(img.naturalHeight * k);
  c.getContext('2d').drawImage(img, 0, 0, c.width, c.height); URL.revokeObjectURL(img.src);
  return await new Promise(res => c.toBlob(res, 'image/jpeg', 0.8));
}
$('sdF').onchange = async e => {
  const f = e.target.files[0]; e.target.value = ''; if (!f || !det) return;
  if (!navigator.onLine) { toast(tr('Las fotos necesitan conexión')); return; }
  const { v, s } = det; toast(tr('Subiendo foto…'));
  try {
    const blob = await shrink(f), path = slug(curId) + '/' + v + '/' + s + '-' + Date.now() + '.jpg';
    await api('/storage/v1/object/fotos/' + path, { raw: true, method: 'POST', blob, timeout: 60000 });
    const url = URL_ + '/storage/v1/object/public/fotos/' + path;
    MARKS[v] = MARKS[v] || {}; MARKS[v][s] = Object.assign({}, MARKS[v][s] || {}, { photo: url });
    queue(v, s, { photo: url, done_by: MARKS[v][s].by || ME || null });
    if (det && det.v === v && det.s === s) openDetail(v, s);
    toast(tr('Foto guardada'));
  } catch (err) { toast(tr('No se pudo subir la foto')); }
};
$('sdH').onclick = async () => {
  if (!det) return; const { v } = det;
  try {
    const rows = await api('events?list_id=eq.' + enc(curId) + '&vin=eq.' + enc(v) + '&select=stage,action,by_name,note,at&order=at.desc&limit=200');
    const act = { done: 'marcó', undo: 'desmarcó', note: 'nota', photo: 'foto' };
    $('hiT').textContent = tr('Historial') + ' · ' + v.slice(-6);
    $('hiL').innerHTML = (rows || []).length ? rows.map(r => '<li><b>' + esc(hm(ms(r.at))) + '</b> · ' + esc(r.by_name || '—') + ' ' + esc(tr(act[r.action] || r.action)) + ' <b>' + esc(stageLabel(r.stage)) + '</b>'
      + (r.action === 'note' && r.note ? '<br>“' + esc(r.note) + '”' : '') + (r.action === 'photo' && /^https:\/\//.test(r.note || '') ? ' <a href="' + esc(r.note) + '" target="_blank" rel="noopener">📷</a>' : '') + '</li>').join('') : '<li>' + esc(tr('Sin cambios todavía.')) + '</li>';
    open_('hi');
  } catch (e) { toast(tr('El historial necesita conexión')); }
};
$('hiC').onclick = () => close_('hi');

/* ---------- copiar y exportar ---------- */
$('copy').onclick = async () => {
  const l = L(); if (!l) return; const S = STAGES();
  const lines = [l.name + ' — ' + new Date().toLocaleString(LOC()), tr('Total') + ': ' + VEH.length].concat(S.map(s => tr(s.label) + ': ' + VEH.filter(x => done(x.vin, s.id)).length), [tr('Terminados (todas las fases)') + ': ' + VEH.filter(x => finished(x.vin)).length]);
  try { await navigator.clipboard.writeText(lines.join('\n')); toast(tr('Copiar resumen') + ' ✓'); } catch (e) { toast(tr('No se pudo copiar')); }
};
let xlsxP = null;
function loadXLSX() { if (window.XLSX) return Promise.resolve(window.XLSX); if (!xlsxP) xlsxP = loadScript('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js').then(() => window.XLSX).catch(e => { xlsxP = null; throw e; }); return xlsxP; }
async function saveFile(name, buf) {
  const file = new File([buf], name, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  if (navigator.canShare && navigator.canShare({ files: [file] })) { try { await navigator.share({ files: [file], title: name }); return true; } catch (e) { if (e && e.name === 'AbortError') return false; } }
  const a = document.createElement('a'); a.href = URL.createObjectURL(file); a.download = name; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000); return true;
}
$('xls').onclick = async () => {
  const l = L(); if (!l) return; toast(tr('Preparando Excel…'));
  let X; try { X = await loadXLSX(); } catch (e) { toast(tr('No se pudo cargar el lector de Excel; revisa la conexión')); return; }
  const S = STAGES(), C = COLS().filter(c => c.show !== 'no');
  const f = t => { const d = new Date(t); return [d.toLocaleDateString('es-AU'), d.toLocaleTimeString('es-AU', { hour:'2-digit', minute:'2-digit' })]; };
  const H = ['VIN'].concat(C.map(c => tr(c.label)));
  S.forEach(s => H.push(tr(s.label) + ' ' + tr('fecha'), tr(s.label) + ' ' + tr('hora'), tr(s.label) + ' ' + tr('por'), tr(s.label) + ' ' + tr('nota')));
  const rows = VEH.map(x => { const r = [x.vin].concat(C.map(c => x[c.key] == null ? '' : x[c.key])); S.forEach(s => { const m = MARKS[x.vin] && MARKS[x.vin][s.id] || {}; r.push(...(m.t ? f(m.t) : ['', '']), m.t ? (m.by || '') : '', m.note || ''); }); return r; });
  const wb = X.utils.book_new(), ws = X.utils.aoa_to_sheet([H].concat(rows));
  ws['!cols'] = H.map((h, i) => ({ wch: i === 0 ? 20 : Math.min(24, Math.max(9, String(h).length + 2)) })); ws['!autofilter'] = { ref: 'A1:' + X.utils.encode_col(H.length - 1) + (rows.length + 1) };
  X.utils.book_append_sheet(wb, ws, tr('Carros'));
  const rs = [[tr('Lista'), l.name], [tr('Exportado'), f(Date.now()).join(' ')], [tr('Total'), VEH.length]].concat(S.map(s => [tr(s.label), VEH.filter(x => done(x.vin, s.id)).length]), [[tr('Terminados (todas las fases)'), VEH.filter(x => finished(x.vin)).length]]);
  const wr = X.utils.aoa_to_sheet(rs); wr['!cols'] = [{ wch: 26 }, { wch: 24 }]; X.utils.book_append_sheet(wb, wr, tr('Resumen'));
  try {
    const ev = await api('events?list_id=eq.' + enc(l.id) + '&select=vin,stage,action,by_name,note,at&order=at.asc&limit=50000', { timeout: 30000 });
    const act = { done: 'marcó', undo: 'desmarcó', note: 'nota', photo: 'foto' };
    const he = X.utils.aoa_to_sheet([['VIN', tr('Fase'), tr('Acción'), tr('Fecha'), tr('Hora'), tr('Persona'), tr('Nota')]].concat((ev || []).map(r => [r.vin, stageLabel(r.stage), tr(act[r.action] || r.action), ...f(ms(r.at)), r.by_name || '', r.note || ''])));
    he['!cols'] = [{ wch: 20 }, { wch: 16 }, { wch: 10 }, { wch: 11 }, { wch: 7 }, { wch: 14 }, { wch: 40 }]; X.utils.book_append_sheet(wb, he, tr('Historial'));
  } catch (e) { /* sin conexión: se exporta sin historial */ }
  const buf = X.write(wb, { bookType: 'xlsx', type: 'array' }), d = new Date();
  const name = 'SalidaVIN_' + l.name.replace(/[^A-Za-z0-9]+/g, '_') + '_' + d.getFullYear() + String(d.getMonth() + 1).padStart(2, '0') + String(d.getDate()).padStart(2, '0') + '.xlsx';
  try { if (await saveFile(name, buf)) toast(tr('Excel listo')); } catch (e) { toast(tr('No se pudo descargar el Excel')); }
};

/* ---------- admin ---------- */
function adminUI() { $('admT').hidden = !ADMIN; $('adm').hidden = !CFG.ADMIN_PIN; $('adm').textContent = tr(ADMIN ? '🔓 Salir de admin' : '🔒 Admin'); }
$('adm').onclick = () => { if (ADMIN) { ADMIN = false; lset('admin', false); adminUI(); return; } $('pinI').value = ''; open_('pin'); setTimeout(() => $('pinI').focus(), 50); };
$('pinN').onclick = () => close_('pin');
$('pinY').onclick = () => { if ($('pinI').value === String(CFG.ADMIN_PIN)) { ADMIN = true; lset('admin', true); close_('pin'); adminUI(); toast(tr('Modo admin activado')); } else toast(tr('PIN incorrecto')); };
$('pinI').addEventListener('keydown', e => { if (e.key === 'Enter') $('pinY').click(); });

/* Importar Excel con columnas libres */
const colLetter = i => { let s = ''; i++; while (i) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
const cleanCell = v => String(v == null ? '' : v).replace(/\s+/g, ' ').trim();
const cleanVin = v => String(v == null ? '' : v).toUpperCase().replace(/[^A-Z0-9]/g, '');
function guessLabel(vals, used) {
  const nz = vals.filter(Boolean), all = re => nz.length && nz.filter(v => re.test(v)).length >= nz.length * 0.8;
  const pick = l => used.has(l) ? '' : l;
  if (all(/^[A-Z]{1,2}\d{1,3}(-\d{1,3}){1,3}$/i)) return pick('Ubicación');
  if (all(/\b(BYD|SEALION|SEAL|ATTO|DOLPHIN|SHARK|TANG|HAN|MODEL)\b/i)) return pick('Modelo');
  if (all(/\b(BLACK|WHITE|BLUE|GREY|GRAY|SILVER|RED|GREEN|PINK|PURPLE|GOLD|BEIGE)\b/i)) return pick('Color');
  if (new Set(nz).size <= 12 && all(/[A-Za-z]/) && !used.has('Destino')) return 'Destino';
  return '';
}
function isCheckCol(vals) { const nz = vals.filter(Boolean); return !nz.length || nz.every(v => /^(TRUE|FALSE|VERDADERO|FALSO|X|✓|✔|☐|☑|0|1|YES|NO|SI|SÍ)$/i.test(v)); }
let IMP = null;
$('file').onchange = async e => {
  const f = e.target.files[0]; e.target.value = ''; if (!f) return;
  try {
    let sheets;
    if (/\.csv$/i.test(f.name)) sheets = [{ name: 'CSV', rows: (await f.text()).split(/\r?\n/).map(l => l.split(/[,;\t]/)) }];
    else { const X = await loadXLSX(); const wb = X.read(await f.arrayBuffer()); sheets = wb.SheetNames.map(n => ({ name: n, rows: X.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: false, defval: '' }) })); }
    let best = null;
    for (const sh of sheets) {
      const width = sh.rows.reduce((w, r) => Math.max(w, r.length), 0), hits = new Array(width).fill(0);
      sh.rows.forEach(r => r.forEach((c, i) => { if (VIN_RE.test(cleanVin(c))) hits[i]++; }));
      const vc = hits.indexOf(Math.max(0, ...hits));
      if (vc >= 0 && hits[vc] > 0 && (!best || hits[vc] > best.n)) best = { sh, vc, n: hits[vc], width };
    }
    if (!best) { toast(tr('No encontré ningún VIN en ese archivo')); return; }
    const { sh, vc, width } = best, data = [], seen = new Set(); let first = -1;
    sh.rows.forEach((r, i) => { const v = cleanVin(r[vc]); if (VIN_RE.test(v) && !seen.has(v)) { seen.add(v); data.push(r); if (first < 0) first = i; } });
    const hdr = first > 0 ? sh.rows[first - 1] : null;
    const hasHdr = !!(hdr && !VIN_RE.test(cleanVin(hdr[vc])) && hdr.filter(c => cleanCell(c)).length >= 2);
    const sig = (hasHdr ? 'h:' + hdr.map(cleanCell).join('|') : 'n:' + width + ':' + vc).toLowerCase();
    const saved = (lget('c:colmaps', {})[sig]) || (TPL.__colmaps && TPL.__colmaps[sig]) || null;
    const used = new Set(), cols = [];
    for (let c = 0; c < width; c++) {
      if (c === vc) continue;
      const vals = data.map(r => cleanCell(r[c]));
      if (!vals.some(Boolean)) continue;
      const head = hasHdr ? cleanCell(hdr[c]) : '', chk = isCheckCol(vals);
      let label = head || (chk ? tr('Casilla') + ' ' + colLetter(c) : guessLabel(vals, used)) || (tr('Columna') + ' ' + colLetter(c)); used.add(label);
      let show = chk ? 'no' : /^(Destino|Modelo|Ubicación)$/.test(label) || /dest|model|loc|bay/i.test(head) ? 'card' : 'detail';
      if (saved && saved[c]) { label = saved[c].label || label; show = saved[c].show || show; }
      cols.push({ c, label, show, sample: vals.find(Boolean) || '' });
    }
    IMP = { file: f.name, sig, vc, data, cols };
    $('impInfo').textContent = data.length + ' ' + tr('vehículos en la hoja') + ' "' + sh.name + '" · ' + tr('VIN en la columna') + ' ' + colLetter(vc)
      + (curId ? ' · ' + data.filter(r => !VEH.some(x => x.vin === cleanVin(r[vc]))).length + ' ' + tr('no están en la lista actual') : '');
    $('impMap').innerHTML = cols.map((k, i) => '<div class="mapr"><label style="gap:2px"><input data-ml="' + i + '" value="' + esc(k.label) + '" maxlength="30" aria-label="' + esc(tr('Columna')) + ' ' + colLetter(k.c) + '"><small>' + colLetter(k.c) + ' · ' + esc(k.sample) + '</small></label>'
      + '<select data-ms="' + i + '" aria-label="' + esc(k.label) + '"><option value="card">' + esc(tr('Ver en tarjeta')) + '</option><option value="detail">' + esc(tr('Solo en detalle')) + '</option><option value="no">' + esc(tr('No guardar')) + '</option></select></div>').join('');
    cols.forEach((k, i) => { document.querySelector('[data-ms="' + i + '"]').value = k.show; });
    $('impName').value = f.name.replace(/\.[^.]+$/, '').slice(0, 40);
    try { const t = await api('settings?key=eq.templates&select=value'); if (t && t[0] && t[0].value) { TPL = t[0].value; lset('c:tpl', TPL); } } catch (e) {}
    const types = Object.keys(TPL).filter(k => !k.startsWith('__')); if (!types.length) types.push('Despacho');
    $('impType').innerHTML = types.map(t => '<option value="' + esc(t) + '">' + esc(tr(t)) + '</option>').join('');
    if (types.includes('Despacho')) $('impType').value = 'Despacho';
    $('impA').hidden = !curId; open_('imp');
  } catch (err) { toast(tr('No se pudo leer el archivo')); }
};
function readMapping() {
  const out = [], keys = new Set(['vin']);
  IMP.cols.forEach((k, i) => {
    const label = (document.querySelector('[data-ml="' + i + '"]').value || '').trim() || (tr('Columna') + ' ' + colLetter(k.c));
    const show = document.querySelector('[data-ms="' + i + '"]').value; k.label = label; k.show = show;
    if (show === 'no') return;
    let key = slug(label), n = 2; while (keys.has(key)) key = slug(label) + '_' + n++; keys.add(key);
    out.push({ c: k.c, key, label, show });
  });
  const maps = lget('c:colmaps', {}); maps[IMP.sig] = {}; IMP.cols.forEach(k => { maps[IMP.sig][k.c] = { label: k.label, show: k.show }; }); lset('c:colmaps', maps);
  return out;
}
function rowsToVehicles(map) { return IMP.data.map(r => { const o = { vin: cleanVin(r[IMP.vc]) }; map.forEach(m => { const v = cleanCell(r[m.c]); if (v) o[m.key] = v; }); return o; }); }
$('impC').onclick = () => { close_('imp'); IMP = null; };
$('impY').onclick = async () => {
  if (!IMP) return; const map = readMapping(), name = $('impName').value.trim() || tr('Lista nueva'), type = $('impType').value;
  const stages = JSON.parse(JSON.stringify(TPL[type] || DEF_TPL[type] || DEF_TPL.Despacho));
  const columns = map.map(m => ({ key: m.key, label: m.label, show: m.show })), vehicles = rowsToVehicles(map);
  const id = slug(name).replace(/_/g, '-') + '-' + Date.now().toString(36);
  close_('imp');
  try { await api('lists', { method: 'POST', prefer: 'return=minimal', body: { id, name, type, columns, stages, vehicles } }); IMP = null; await pullLists(); await openList(id); toast(tr('Lista creada') + ': ' + vehicles.length); }
  catch (e) { toast(tr('No se pudo guardar; revisa la conexión')); }
};
$('impA').onclick = async () => {
  if (!IMP || !curId) return; const map = readMapping(), l = L(); close_('imp');
  try {
    const rows = await api('lists?id=eq.' + enc(curId) + '&select=vehicles,columns');
    const base = rows && rows[0] && Array.isArray(rows[0].vehicles) ? rows[0].vehicles : VEH, columns = rows && rows[0] && Array.isArray(rows[0].columns) ? rows[0].columns.slice() : COLS().slice();
    map.forEach(m => { const ex = columns.find(c => (c.label || '').toLowerCase() === m.label.toLowerCase()); if (ex) m.key = ex.key; else { let k = m.key, n = 2; while (columns.some(c => c.key === k)) k = m.key + '_' + n++; m.key = k; columns.push({ key: k, label: m.label, show: m.show }); } });
    const have = new Set(base.map(x => x.vin)), add = rowsToVehicles(map).filter(x => !have.has(x.vin));
    await api('lists?id=eq.' + enc(curId), { method: 'PATCH', prefer: 'return=minimal', body: { vehicles: base.concat(add), columns } });
    IMP = null; vehUpdated = null; await pullLists(); await pullVehicles(); toast(add.length + ' ' + tr('vehículos agregados') + ' · ' + (l ? l.name : ''));
  } catch (e) { toast(tr('No se pudo guardar; revisa la conexión')); }
};

/* Editor de fases */
let FE = [];
function feRender() {
  $('feList').innerHTML = FE.map((s, i) => s.hidden ? '' : '<div class="fe"><input type="color" data-fc="' + i + '" value="' + okColor(s.color) + '" aria-label="Color"><input data-fl="' + i + '" value="' + esc(s.label) + '" maxlength="24" aria-label="' + esc(tr('Fase')) + '">'
    + '<button data-fu="' + i + '" aria-label="' + esc(tr('Subir')) + '">↑</button><button data-fd="' + i + '" aria-label="' + esc(tr('Bajar')) + '">↓</button><button data-fx="' + i + '" aria-label="' + esc(tr('Quitar')) + '">✕</button></div>').join('');
}
function feRead() { FE.forEach((s, i) => { const l = document.querySelector('[data-fl="' + i + '"]'), c = document.querySelector('[data-fc="' + i + '"]'); if (l) s.label = l.value.trim() || s.label; if (c) s.color = c.value; }); }
$('fases').onclick = () => { const l = L(); if (!l) return; FE = JSON.parse(JSON.stringify(l.stages || [])); $('feTpl').checked = false; feRender(); open_('fe'); };
$('feList').addEventListener('click', e => {
  const b = e.target.closest('button'); if (!b) return; feRead(); const d = b.dataset, vis = FE.map((s, i) => s.hidden ? -1 : i).filter(i => i >= 0);
  const mv = (i, dir) => { const p = vis.indexOf(i), j = vis[p + dir]; if (j === undefined) return; [FE[i], FE[j]] = [FE[j], FE[i]]; };
  if (d.fu) mv(+d.fu, -1); else if (d.fd) mv(+d.fd, 1); else if (d.fx) FE[+d.fx].hidden = true;
  feRender();
});
$('feAdd').onclick = () => { feRead(); let id = 'fase', n = 2; while (FE.some(s => s.id === id)) id = 'fase' + n++; FE.push({ id, label: tr('Fase nueva') + ' ' + (FE.filter(s => !s.hidden).length + 1), color: '#c2410c' }); feRender(); };
$('feC').onclick = () => close_('fe');
$('feY').onclick = async () => {
  feRead(); const l = L(); if (!l) return;
  FE.forEach(s => { if (/^fase\d*$/.test(s.id)) { let id = slug(s.label), n = 2; while (FE.some(o => o !== s && o.id === id)) id = slug(s.label) + n++; s.id = id; } });
  try {
    await api('lists?id=eq.' + enc(l.id), { method: 'PATCH', prefer: 'return=minimal', body: { stages: FE } });
    if ($('feTpl').checked && l.type) { const t = Object.assign({}, TPL); t[l.type] = FE.filter(s => !s.hidden).map(s => ({ id: s.id, label: s.label, color: s.color })); await api('settings?on_conflict=key', { method: 'POST', prefer: 'resolution=merge-duplicates,return=minimal', body: [{ key: 'templates', value: t }] }); }
    close_('fe'); await pullLists(); toast(tr('Fases guardadas'));
  } catch (e) { toast(tr('No se pudo guardar; revisa la conexión')); }
};

/* Editor de columnas */
let CE = [];
$('cols').onclick = () => { const l = L(); if (!l) return; CE = JSON.parse(JSON.stringify(l.columns || []));
  $('ceList').innerHTML = CE.map((c, i) => '<div class="mapr"><input data-cl="' + i + '" value="' + esc(c.label) + '" maxlength="30" aria-label="' + esc(tr('Columna')) + '"><select data-cs="' + i + '" aria-label="' + esc(c.label) + '"><option value="card">' + esc(tr('Ver en tarjeta')) + '</option><option value="detail">' + esc(tr('Solo en detalle')) + '</option><option value="no">' + esc(tr('No guardar')) + '</option></select></div>').join('');
  CE.forEach((c, i) => { document.querySelector('[data-cs="' + i + '"]').value = c.show || 'detail'; }); open_('ce'); };
$('ceC').onclick = () => close_('ce');
$('ceY').onclick = async () => { const l = L(); if (!l) return;
  CE.forEach((c, i) => { c.label = document.querySelector('[data-cl="' + i + '"]').value.trim() || c.label; c.show = document.querySelector('[data-cs="' + i + '"]').value; });
  try { await api('lists?id=eq.' + enc(l.id), { method: 'PATCH', prefer: 'return=minimal', body: { columns: CE } }); close_('ce'); await pullLists(); toast(tr('Columnas guardadas')); } catch (e) { toast(tr('No se pudo guardar; revisa la conexión')); } };

$('reset').onclick = () => { const l = L(); if (!l) return;
  ask(tr('¿Desmarcar TODAS las fases de todos los carros de') + ' "' + l.name + '" ' + tr('para todos? Los carros y el historial se mantienen.'), async () => {
    try { await api('marks?list_id=eq.' + enc(l.id), { method: 'PATCH', prefer: 'return=minimal', body: { done_at: null, done_by: ME || null } });
      OUTBOX = OUTBOX.filter(x => x.list_id !== l.id); lset('outbox2', OUTBOX); rebuildPending(); await pullMarks(true); toast(tr('Lista reiniciada')); }
    catch (e) { toast(tr('No se pudo guardar; revisa la conexión')); } }); };
$('arch').onclick = () => { const l = L(); if (!l) return;
  const go = async () => { try { await api('lists?id=eq.' + enc(l.id), { method: 'PATCH', prefer: 'return=minimal', body: { archived: !l.archived } }); await pullLists(); toast(tr(l.archived ? 'Lista activa otra vez' : 'Lista archivada')); } catch (e) { toast(tr('No se pudo guardar; revisa la conexión')); } };
  if (l.archived) go(); else ask(tr('¿Archivar') + ' "' + l.name + '"' + tr('? Sigue guardada y puedes consultarla en "Archivadas".'), go); };

/* ---------- escáner ---------- */
let scan = null, lastCode = '', lastAt = 0;
function beep() { try { const a = new (window.AudioContext || window.webkitAudioContext)(), o = a.createOscillator(), g = a.createGain(); o.frequency.value = 880; g.gain.value = 0.15; o.connect(g); g.connect(a.destination); o.start(); setTimeout(() => { o.stop(); a.close(); }, 120); } catch (e) {} if (navigator.vibrate) navigator.vibrate(60); }
function vinFrom(text) {
  const up = String(text || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  const exact = VEH.find(x => up.includes(x.vin)); if (exact) return { x: exact, vin: exact.vin };
  const m = up.match(/[A-HJ-NPR-Z0-9]{17}/g) || []; const v = m.length ? (up.length === 18 && up[0] === 'I' ? up.slice(1) : m[0]) : '';
  return { x: VEH.find(x => x.vin === v) || null, vin: v || up };
}
function onCode(text) {
  const now = Date.now(); if (text === lastCode && now - lastAt < 2500) return; lastCode = text; lastAt = now;
  const { x, vin } = vinFrom(text);
  if (!x) { $('camS').textContent = tr('No está en esta lista') + ': ' + String(vin).slice(0, 40); return; }
  beep();
  if (quick) { quickMark(x); const s = STAGES().find(z => z.id === quick); $('camS').textContent = '⚡ ' + x.vin.slice(-6) + ' · ' + (s ? tr(s.label) : '') + ' ✓'; return; }
  stopScan(); close_('cm'); $('q').value = x.vin; search(); $('res').scrollIntoView({ behavior: 'smooth', block: 'center' });
}
const CAMV = { facingMode: 'environment', width: { ideal: 1280 }, height: { ideal: 720 }, advanced: [{ focusMode: 'continuous' }] };
async function startScan() {
  const v = $('camv'); $('camS').textContent = tr('Abriendo cámara…'); open_('cm');
  try {
    if ('BarcodeDetector' in window) {
      const fm = await window.BarcodeDetector.getSupportedFormats().catch(() => []);
      const want = ['code_39', 'code_128', 'qr_code', 'data_matrix', 'pdf417'].filter(f => fm.includes(f));
      if (want.length) {
        const det = new window.BarcodeDetector({ formats: want });
        const stream = await navigator.mediaDevices.getUserMedia({ video: CAMV, audio: false });
        v.srcObject = stream; await v.play(); $('camS').textContent = tr('Apunta al código QR o de barras del VIN');
        let alive = true; scan = { stop: () => { alive = false; stream.getTracks().forEach(t => t.stop()); v.srcObject = null; } };
        const loop = async () => { if (!alive) return; try { const r = await det.detect(v); if (r && r[0]) onCode(r[0].rawValue); } catch (e) {} setTimeout(loop, 200); };
        loop(); return;
      }
    }
    if (!window.ZXing) await loadScript('https://cdn.jsdelivr.net/npm/@zxing/library@0.21.3/umd/index.min.js');
    const Z = window.ZXing, hints = new Map();
    hints.set(Z.DecodeHintType.POSSIBLE_FORMATS, [Z.BarcodeFormat.CODE_39, Z.BarcodeFormat.CODE_128, Z.BarcodeFormat.QR_CODE, Z.BarcodeFormat.DATA_MATRIX, Z.BarcodeFormat.PDF_417]);
    hints.set(Z.DecodeHintType.TRY_HARDER, true);
    const reader = new Z.BrowserMultiFormatReader(hints, 300);
    scan = { stop: () => { try { reader.reset(); } catch (e) {} } };
    const cb = r => { if (r) onCode(r.getText ? r.getText() : r.text); };
    if (reader.decodeFromConstraints) await reader.decodeFromConstraints({ video: CAMV }, v, cb);
    else await reader.decodeFromVideoDevice(undefined, v, cb);
    $('camS').textContent = tr('Apunta al código QR o de barras del VIN');
  } catch (e) { $('camS').textContent = tr('No se pudo abrir la cámara. Revisa el permiso en Ajustes > Safari > Cámara.'); }
}
function stopScan() { if (scan) { try { scan.stop(); } catch (e) {} scan = null; } }
$('camb').onclick = startScan;
$('camC').onclick = () => { stopScan(); close_('cm'); };

/* ---------- nombre, pantalla completa, idioma ---------- */
function whoUI() { $('who').textContent = '👤 ' + (ME || '—') + ' · ' + tr('Cambiar nombre'); }
function askName() { $('nmI').value = ME; open_('nm'); setTimeout(() => $('nmI').focus(), 50); }
$('nmY').onclick = () => { const v = $('nmI').value.trim(); if (!v) { $('nmI').focus(); return; } ME = v; lset('me', ME); close_('nm'); whoUI(); };
$('nmI').addEventListener('keydown', e => { if (e.key === 'Enter') $('nmY').click(); });
$('who').onclick = askName;
(function () {
  const d = document, E = d.documentElement, b = $('fs'), req = E.requestFullscreen || E.webkitRequestFullscreen;
  if (!req || !(d.fullscreenEnabled || d.webkitFullscreenEnabled)) { b.hidden = true; return; }
  const on = () => !!(d.fullscreenElement || d.webkitFullscreenElement);
  window.__fsUpd = () => { b.textContent = tr(on() ? '✕ Salir' : '⛶ Pantalla completa'); };
  b.onclick = async () => { try { if (on()) await (d.exitFullscreen || d.webkitExitFullscreen).call(d); else await req.call(E, { navigationUI: 'hide' }); } catch (e) { toast(tr('Este navegador no permite pantalla completa aquí')); } };
  d.addEventListener('fullscreenchange', window.__fsUpd); d.addEventListener('webkitfullscreenchange', window.__fsUpd);
})();
const TN = [];
(function walk(n) { for (const c of n.childNodes) { if (c.nodeType === 3) { const t = c.nodeValue.trim(); if (t && EN[t] !== undefined) TN.push([c, c.nodeValue, t]); } else if (c.nodeType === 1 && !/^(SCRIPT|STYLE|SELECT|UL)$/.test(c.tagName) && !['res', 'list', 'sync', 'lname', 'who', 'adm', 'fs', 'lng', 'arch', 'seg', 'cards', 'hm', 'cf', 'sum', 'qban', 'impMap', 'impInfo', 'feList', 'ceList', 'sdT', 'sdS', 'sdP', 'hiT', 'camS', 'dlgT'].includes(c.id)) walk(c); } })(document.body);
function applyLang() {
  document.documentElement.lang = lang === 'zh' ? 'zh-Hans' : lang;
  for (const [n, orig, t] of TN) n.nodeValue = lang === 'es' ? orig : orig.replace(t, tr(t));
  $('lng').value = lang; $('lng').setAttribute('aria-label', tr('Idioma'));
  $('clr').setAttribute('aria-label', tr('Borrar')); $('camb').setAttribute('aria-label', tr('Escanear código')); document.title = tr('Salida VIN');
  if (window.__fsUpd) window.__fsUpd(); else $('fs').textContent = tr('⛶ Pantalla completa');
  adminUI(); whoUI(); fillLists(); fillQuick(); fillFilters(); render(); syncUI();
}
$('lng').onchange = () => { lang = $('lng').value; try { localStorage.setItem('lang', lang); } catch (e) {} applyLang(); };

/* ======================= arranque ======================= */
$('ver').textContent = 'v' + VERSION;
$('cfg').hidden = CONFIGURED;
if (curId) loadCache(curId);
applyLang();
if (!ME) askName();
window.addEventListener('online', () => { online = true; poll(); });
window.addEventListener('offline', () => { online = false; syncUI(); });
document.addEventListener('visibilitychange', () => { if (!document.hidden) { tick = 0; poll(); } else stopScan(); });
if (CONFIGURED) { poll(); setInterval(poll, 5000); }  // cambios nuevos cada 5 s
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    const had = !!navigator.serviceWorker.controller;
    navigator.serviceWorker.register('sw.js').then(reg => { reg.addEventListener('updatefound', () => { const w = reg.installing; if (w) w.addEventListener('statechange', () => { if (w.state === 'activated' && had) toast(tr('Hay una versión nueva; se actualizó')); }); }); }).catch(() => {});
  });
}
})();
