unit srv_Comunicacion;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, SvcMgr, Dialogs,
  ExtCtrls, DB, ADODB, Registry, uSucursal, Contnrs, uParametro, Variants, WinSvc,
  ShellAPI, AppEvnts, ImgList, Menus, uParametro_Formato, IdBaseComponent,
  IdComponent, IdTCPConnection, IdTCPClient, IdExplicitTLSClientServerBase,
  IdFTP, IdFTPCommon, uEmpresa, uGestor_Empresas, uGestor_Log_Informe_Automatico, uGestor_Articulos,
  uGestor_Lanzador_Zafiro, uParametro_Zafiro_BI, uGestor_Log_subida_comprobantes, uGestor_Sucursales,
  uGestor_Comprobante_Ventas, uGestor_Comprobante_Stock, dm_Comprobante_Ventas, uNiveles_de_Rotacion, uGestor_Niveles_de_Rotacion, System.Math,
  System.ImageList;

const
  WM_ICONTRAY = WM_USER + 1;

type

  TArrInteger = array of Integer;

  TVSComunicacionParaReplicacion01 = class(TService)
    qrySelectAll_sucursales_comunicacion: TADOQuery;
    qrySelect_Binlog_events: TADOQuery;
    qrySelect_LastBinaryLogOrigen: TADOQuery;
    qryUpdate_FileandPos_BDorigen: TADOQuery;
    qrySentencias_comunicacion: TADOQuery;
    Conexion_Remota: TADOConnection;
    qrySelect_tablas_exceptuadas_por_sucursal: TADOQuery;
    qryUpdate_FileandPos_BDDestino: TADOQuery;
    qryDesactivar_log_local: TADOQuery;
    Timer1: TTimer;
    ImageList1: TImageList;
    //CoolTrayIcon1: TCoolTrayIcon;
    Conexion: TADOConnection;
    PopupMenu_TrayServicio: TPopupMenu;
    ppmReiniciar_servicio: TMenuItem;
    ppmParar_servicio: TMenuItem;
    N1: TMenuItem;
    StopService1: TMenuItem;
    qrySelect_LastBinaryLogDestino: TADOQuery;
    qryLog_comunicacion_externo: TADOQuery;
    N2: TMenuItem;
    qrySelect_Binlog_events_Destino: TADOQuery;
    Timer2: TTimer;
    Timer3: TTimer;
    Timer4: TTimer;
    Timer5: TTimer;
    Timer6: TTimer;
    Timer7: TTimer;
    Timer8: TTimer;
    Timer9: TTimer;
    Timer10: TTimer;
    Timer11: TTimer;
    Timer12: TTimer;
    Timer13: TTimer;
    Timer14: TTimer;
    Timer15: TTimer;
    Timer16: TTimer;
    Timer17: TTimer;
    Timer18: TTimer;
    Timer19: TTimer;
    Timer20: TTimer;
    Timer100: TTimer;   // para Subir a Amazon AWS

    qryUpdateNormativa: TADOQuery;
    qrySelectLogProcesadoOrigenEnDestino: TADOQuery;
    qryUpdateArticuloGestion: TADOQuery;
    qrySelect_Binlog_events_Destino_EPosterior: TADOQuery;
    qryDescribe_tabla: TADOQuery;
    qrySelectManual: TADOQuery;
    TimerEtiquetas: TTimer;
    qrySelectAuditoria_Articulo_Resumido: TADOQuery;
    qrySelectAuditoria_Articulo_Resumidofecha_hora: TStringField;
    qrySelectAuditoria_Articulo_Resumidoid_articulo: TStringField;
    qrySelectAuditoria_Articulo_Resumidodes_articulo: TStringField;
    qrySelectAuditoria_Articulo_Resumidopresentacion: TStringField;
    qrySelectAuditoria_Articulo_Resumidodes_marca: TStringField;
    qrySelectAuditoria_Articulo_Resumidocod_barra: TStringField;
    qrySelectAuditoria_Articulo_ResumidoPcio_vta_civa: TFloatField;
    qrySelectAuditoria_Articulo_Resumidodes_arti_pres: TStringField;
    TimerEsperaConsultaIQVIA: TTimer;
    TimerExportarRecetasCloseup: TTimer;
    qrySelectExportacion: TADOQuery;
    IdFTP1: TIdFTP;
    TimerActualizacionPreciosProgramada: TTimer;
    TimerActualizacionZafiroBI: TTimer;
    TimerSubidaAutoFTPCompPend: TTimer;
    TimerActualizacion_Min_Max_Stock: TTimer;
    qryUpdate_Comu_Fecha_Hora_Ultima_Revision: TADOQuery;
    qryUpdate_Comu_Fecha_Hora: TADOQuery;
    qrySelect_Now: TADOQuery;
    qrySelect_Comu_Fecha_Hora: TADOQuery;
    qryUpdate_sucursales_log_tiempo_comunicacion_local: TADOQuery;
    Timer21: TTimer;
    Timer22: TTimer;
    Timer23: TTimer;
    Timer24: TTimer;
    Timer25: TTimer;
    Timer26: TTimer;
    Timer27: TTimer;
    Timer28: TTimer;
    TimerActualizacionPCpraPP: TTimer;
    TimerExportacionDataview: TTimer;
    qrySelectTicketsDia: TADOQuery;
    qryTodas_Sucursales: TADOQuery;

    procedure ServiceAfterInstall(Sender: TService);
    procedure ServiceExecute(Sender: TService);
    procedure Timer1Timer(Sender: TObject);
    procedure ppmParar_servicioClick(Sender: TObject);
    procedure ServiceDestroy(Sender: TObject);
    procedure ppmReiniciar_servicioClick(Sender: TObject);
    procedure StopService1Click(Sender: TObject);
    procedure TimerEtiquetasTimer(Sender: TObject);
    procedure TimerEsperaConsultaIQVIATimer(Sender: TObject);
    procedure TimerExportarRecetasCloseupTimer(Sender: TObject);
    procedure TimerActualizacionPreciosProgramadaTimer(Sender: TObject);
    procedure TimerActualizacionZafiroBITimer(Sender: TObject);
    procedure TimerSubidaAutoFTPCompPendTimer(Sender: TObject);
    procedure TimerActualizacion_Min_Max_StockTimer(Sender: TObject);
    procedure ServiceCreate(Sender: TObject);
    procedure TimerActualizacionPCpraPPTimer(Sender: TObject);
    procedure TimerExportacionDataviewTimer(Sender: TObject);


  private

    oSucursales                    : TObjectList;
    sUltimo_archivo_log_BDorigen   : String;
    iUltimo_Pos_log_BDorigen       : Longint;
    oParametro                     : TParametro;
    oParametros_Formato            : TParametro_Fotmato;
    _Gestor_Empresa                : TGestor_Empresa;
    _Gestor_Sucursal               : TGestor_Sucursal;
    _Gestor_Log_Informe_Automatico : TGestor_Log_Informe_Automatico;
    _Gestor_log_subida_Comp        : TGestor_Log_subida_comprobantes;
    slSentencias_comunicacion      : TStringList;
    _Gestor_Articulos              : TGestor_Articulos;
    _Gestor_Niveles_Rotacion       : TGestor_Niveles_de_Rotacion;
    _Gestor_Comprobante_Ventas     : TGestor_Comprobante_Ventas;
    _dm_Comrobante_Ventas          : TdmComprobante_Ventas;


    bErrorComunicacion, bErrorComunicacionConexionLocal: Boolean;
    iSucursal_procesando1: integer;
    iSucursal_procesando2: integer;
    iSucursal_procesando3: integer;
    iSucursal_procesando4: integer;
    iSucursal_procesando5: integer;
    iSucursal_procesando6: integer;
    iSucursal_procesando7: integer;
    iSucursal_procesando8: integer;
    iSucursal_procesando9: integer;
    iSucursal_procesando10: integer;
    iSucursal_procesando11: integer;
    iSucursal_procesando12: integer;
    iSucursal_procesando13: integer;
    iSucursal_procesando14: integer;
    iSucursal_procesando15: integer;
    iSucursal_procesando16: integer;
    iSucursal_procesando17: integer;
    iSucursal_procesando18: integer;
    iSucursal_procesando19: integer;
    iSucursal_procesando20: integer;
    iSucursal_procesando21: integer;
    iSucursal_procesando22: integer;
    iSucursal_procesando23: integer;
    iSucursal_procesando24: integer;
    iSucursal_procesando25: integer;
    iSucursal_procesando26: integer;
    iSucursal_procesando27: integer;
    iSucursal_procesando28: integer;
    iSucursal_procesando100: integer;

    iCantidad_Dias_Consulta_Iqvia: SmallInt;
    dtFecha_Proceso_IQVIa :TDate;

    //dtFecha_Proceso_ActuPreciosProgramada :TDate;
    //

    function _ActualizarBD(pSucursal: TSucursal; pTabla_Exportar: String = ''): Boolean;    // Devuelve si debe seguir procesando la misma sucursal
    procedure _Grabar_LogErrores_Comunicacion(pId_ProcesoActualizacion: Smallint; pCadenaError: AnsiString; pId_Sucursal: Integer=0);
    procedure _Agregar_CadenaSQL(pSucursal: TSucursal; pCadena:String; pPosicion_caracter_final: Integer; pTabla_Exportar: String = '');
    function _Extraer_NombreTabla(pCadena:String; pPosicion_caracter_final: Integer): String;

    function _Existe_log_Posterior(pId_archivo_log_procesado:String): Boolean;
    function _Existe_Mas_Sentencias_del_mismo_log(pId_archivo_log_procesado:String; pUltimo_Pos : Longint): Boolean;
    function _Carga_Sucursales(pId_ProcesoActualizacion: Smallint; pErrorComunicacion: Boolean):TObjectList;
    function _Actualiza_Ultimo_Log_Procesado(pId_ProcesoActualizacion: Smallint):Boolean;
    function _ServiceStop(pMachine, pServiceName: String): Boolean;
    procedure _ServiceStart(pServiceName: String);
    procedure _CargarItemsMenu;


    function _Get_Pos_log_procesado_Grabado_Destino(pSentencias_comunicacion: TStringList; pId_archivo_log_procesado_Grabado_Destino: String; pPos_log_procesado_Grabado_Destino: Longint; pSentencias_Dep_Texto_Enriquecido: TArrInteger) : Longint;
    function _Existe_log_Posterior_Destino(pId_archivo_log_procesado:String): Boolean;

    procedure _Hab_des_Timer(pTag: Integer; bValor: Boolean);

    function _Depurar_Texto_Erriquecido(pTexto: WideString): WideString;
    function CorrijeCadenaParticular(pCadena : String ): String;
    Function _Confecciona_Sentencias_Faltantes(pTipo_Comprobante: String; pParamatro1: String; pParamatro2: String; pParamatro3: String; pParamatro4: String; pSucursal: TSucursal): TStringList;
    Function _Confecciona_Linea_Sentencias_Faltantes(pTabla: String; pSentencia: String; var pListaSentencias: TStringList; pSucursal: TSucursal; pTipo_Sentencia: String ='INSERT'): Boolean;
    Function _Tabla_Exceptuada(pTabla: String;  pSucursal: TSucursal): Boolean;

    procedure _Consultar_Cupones_No_Informados_IQVIA(pModo: String; pFecha_Desde: TDate; pFecha_Hasta: TDate);

    procedure _Informar_IQVIA_externo(pId_Empresa_Archivo: Integer; pId_Sucursal_Archivo: Integer; pUbicacion_archivos_origen: String; pUbicacion_archivos_procesados: String);

    //Metodos de CloseUp
    procedure _Conectar_Servidor_FTP(pHost, pUser, pPass: String; pPort, pTimeout: Integer; pModoPasivo : Boolean = False);
    procedure _Desconectar_Servidor_FTP;
    function  _Cargar_Archivo_En_Servidor_FTP(pRutaLocal, pArchivo : String; pModo: String = ''): Boolean;
    function  _Descargar_Archivo_De_Servidor_FTP(pRutaLocal, pArchivo: String): Boolean;
    function  _Generar_Lista_Archivos_Remotos: TStringList;
    procedure _Exportar_Datos_Receta;
    procedure _Exportacion_DataView;
    function  _Archivo_Cargado_FTP(pArchivo: string; pLista: TStringList): Boolean;

    function _Actualizar_Precios_Programados: String;

    procedure _Subir_Comp_Pend_Servidor_FTP;

    procedure _Actualizar_Min_Max_Stock(pId_Empresa:Integer);
    function _ObtenerNivele_Rotacion(pId_Empresa: Integer; pId_sucursal: Integer; pId_articulo: String; pFecha_desde: TDate; pFecha_hasta: TDate):TNiveles_de_Rotacion;

    function _Actualizar_Zafiro_BI(pParametro_Zafiro_BI: TParametro_Zafiro_BI; pFecha_Hasta: TDate): String;

    function _Modificar_Comu_Fecha_Hora_Ultima_Revision(pId_Sucursal: Integer): String;
    function _Transmision_de_Sentencias_Por_Comprobante(pSucursal: TSucursal): String;

    function _Ejecuta_Transmision_de_Sentencias_Por_Comprobante(pSucursal: TSucursal): String;

    function _Modificar_Comu_Fecha_Hora(pId_Sucursal: Integer): String;



    //function _Log_Procesado_En_Destino_Es_Mayor_Que_Local(pId_sucursal: Integer; pId_archivo_log_procesado:String; pUltimo_Pos : Longint): Boolean;

    { Private declarations }
  public

    function GetServiceController: TServiceController; override;
    { Public declarations }
  end;

var
  VSComunicacionParaReplicacion01: TVSComunicacionParaReplicacion01;

implementation

uses StrUtils, _utils, DateUtils, frm_AcercaDe, uInterfaceArchivos, uIQVIA,
  uComprobante_Ventas, uTipo_Validacion_Receta, uGestor_Tipo_Validacion_Receta, uComprobante_Stock,
  uLog_Informe_Automatico, uArticulo, uGestor_Parametros, uGestor_Entidades, Datasnap.DBClient, ComCtrls, uGestor_Comunicacion_Manual,
  IdHTTP, IdException, IdStack, System.Zip, System.Net.HttpClient, System.JSON,
  System.Net.URLClient, System.Net.HttpClientComponent, MahExcel;


{$R *.DFM}

procedure ServiceController(CtrlCode: DWord); stdcall;
begin
  VSComunicacionParaReplicacion01.Controller(CtrlCode);
end;

function TVSComunicacionParaReplicacion01.GetServiceController: TServiceController;
begin
  Result := ServiceController;
end;

procedure TVSComunicacionParaReplicacion01.ServiceExecute(Sender: TService);

var
  dsSucursales: TDataSet;
  oSucursal : TSucursal;
  bHaySucursales: Boolean;
  wDia, wMes, wAnio, wHora, wMin, wSeg, wMil : Word;
  dtFechaActual : TDateTime;
begin
//  CoolTrayIcon1.Hint:='V&S Servicio de Comunicación. '+IntToStr(Self.Tag);
  if not Assigned(oParametro) then
  begin
    try
      oParametro:= TParametro.Create;
      //_Grabar_LogErrores_Comunicacion(TTimer(Sender).Tag, 'Creo Parametros');
    except
      on E: Exception do
      begin
        _Grabar_LogErrores_Comunicacion(TTimer(Sender).Tag, 'No se pueden obtener Parametros. Error: '+E.Message);
      end;
    end;

    try
      oParametros_Formato := TParametro_Fotmato.Create;
    except
      on E: Exception do
      begin
        _Grabar_LogErrores_Comunicacion(TTimer(Sender).Tag, 'No se pueden obtener Parametros Formato. Error: '+E.Message);
      end;
    end;

    try
      _Gestor_Empresa := TGestor_Empresa.Create;
    except
      on E: Exception do
      begin
        _Grabar_LogErrores_Comunicacion(TTimer(Sender).Tag, 'No se puede crear _Gestor_Empresa. Error: '+E.Message);
      end;
    end;

    try
      _Gestor_Sucursal := TGestor_Sucursal.Create;
    except
      on E: Exception do
      begin
        _Grabar_LogErrores_Comunicacion(TTimer(Sender).Tag, 'No se puede crear _Gestor_Sucursal. Error: '+E.Message);
      end;
    end;


    try
      _Gestor_Log_Informe_Automatico := TGestor_Log_Informe_Automatico.Create;
    except
      on E: Exception do
      begin
        _Grabar_LogErrores_Comunicacion(TTimer(Sender).Tag, 'No se puede crear _Gestor_Log_Informe_Automatico. Error: '+E.Message);
      end;
    end;

    try
      _Gestor_log_subida_Comp := TGestor_Log_subida_comprobantes.Create;
    except
      on E: Exception do
      begin
        _Grabar_LogErrores_Comunicacion(TTimer(Sender).Tag, 'No se puede crear _Gestor_log_subida_Comp. Error: '+E.Message);
      end;
    end;

    try
      _Gestor_Comprobante_Ventas := TGestor_Comprobante_Ventas.Create;
    except
      on E: Exception do
      begin
        _Grabar_LogErrores_Comunicacion(TTimer(Sender).Tag, 'No se puede crear _Gestor_Comprobante_Ventas. Error: '+E.Message);
      end;
    end;

    try
      _dm_Comrobante_Ventas := TdmComprobante_Ventas.Create(self);
    except
      on E: Exception do
      begin
        _Grabar_LogErrores_Comunicacion(TTimer(Sender).Tag, 'No se puede crear _dm_Comrobante_Ventas. Error: '+E.Message);
      end;
    end;

    TimerEtiquetas.Interval:= 0;
    TimerEsperaConsultaIQVIA.Interval:= 0;
    TimerExportarRecetasCloseup.Interval:= 0;
    TimerExportacionDataview.Interval:= 0;
    TimerActualizacionPreciosProgramada.Interval:= 0;
    TimerActualizacionZafiroBI.Interval:= 0;
    TimerSubidaAutoFTPCompPend.Interval:= 0;
    TimerActualizacion_Min_Max_Stock.Interval := 0;
    TimerActualizacionPCpraPP.Interval := 0;
    //dtFechaHoraUltimaRegistroLog:= IncHour(Now(), -1);

    _Gestor_Articulos              := Nil;
    _Gestor_Niveles_Rotacion       := Nil;

    if Assigned(oParametro) then
    begin
      if oParametro.Comu_Intervalo_Timer > 0 then
      begin
        Timer1.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer2.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer3.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer4.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer5.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer6.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer7.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer8.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer9.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer10.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer11.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer12.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer13.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer14.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer15.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer16.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer17.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer18.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer19.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer20.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer21.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer22.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer23.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer24.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer25.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer26.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer27.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer28.Interval:= oParametro.Comu_Intervalo_Timer;
        Timer100.Interval:= oParametro.Comu_Intervalo_Timer;


      end
      else
      begin
        Timer1.Interval:= 1000;
        Timer2.Interval:= 1000;
        Timer3.Interval:= 1000;
        Timer4.Interval:= 1000;
        Timer5.Interval:= 1000;
        Timer6.Interval:= 1000;
        Timer7.Interval:= 1000;
        Timer8.Interval:= 1000;
        Timer9.Interval:= 1000;
        Timer10.Interval:= 1000;
        Timer11.Interval:= 1000;
        Timer12.Interval:= 1000;
        Timer13.Interval:= 1000;
        Timer14.Interval:= 1000;
        Timer15.Interval:= 1000;
        Timer16.Interval:= 1000;
        Timer17.Interval:= 1000;
        Timer18.Interval:= 1000;
        Timer19.Interval:= 1000;
        Timer20.Interval:= 1000;
        Timer21.Interval:= 1000;
        Timer22.Interval:= 1000;
        Timer23.Interval:= 1000;
        Timer24.Interval:= 1000;
        Timer25.Interval:= 1000;
        Timer26.Interval:= 1000;
        Timer27.Interval:= 1000;
        Timer28.Interval:= 1000;
        Timer100.Interval:= 1000;
      end;

      if oParametro.Etiq_elec_cant_dias_periodo > 0 then
        if Length(Trim(oParametro.Etiq_elec_carpeta_destino)) > 0 then
          if DirectoryExists(oParametro.Etiq_elec_carpeta_destino) then
            if oParametro.Etiq_elec_periodicidad_minutos > 0 then
              TimerEtiquetas.Interval := oParametro.Etiq_elec_periodicidad_minutos * 1000 * 60;
              //TimerEtiquetas.Interval := 10000

      if Self.Tag=1 then   // Solo en el tag = 1 se habilita el timer. Osea en un solo servicio de tag 1
      begin
        if oParametro.Informar_IQVIA then
        begin
          TimerEsperaConsultaIQVIA.Interval:= 330000;  // establezco el timer cada 5:30 minutos (330000 milisegundos)
          iCantidad_Dias_Consulta_Iqvia:= 30; // ocho días
          dtFecha_Proceso_IQVIa  := IncDay(DateOf(Now),-1);
        end;
      end;

      if oParametro.Informar_CloseUp then
      begin
        if Self.Tag=1 then   // Solo en el tag = 1 se habilita el timer. Osea en un solo servicio de tag 1
        begin
          TimerExportarRecetasCloseup.Interval:= 900000;  // establezco el timer cada 15 minutos
        end;
      end;

      if Assigned(_Gestor_Sucursal) then
        oSucursal := _Gestor_Sucursal._Buscar(oParametro.Id_Empresa, oParametro.Id_Sucursal);
      if Assigned(oSucursal) then
      begin
        if Length(oParametro.DV_URL_Base) > 0 then
        begin
          if Self.Tag=1 then   // Solo en el tag = 1 se habilita el timer. Osea en un solo servicio de tag 1
          begin
            TimerExportacionDataview.Interval:= 600000;  // establezco el timer cada 10 minutos
          end;
        end;
      end;

      if Self.Tag=1 then   // Solo en el tag = 1 se habilita el timer. Osea en un solo servicio de tag 1
      begin
        if oParametro.Prec_Activar_Actualizacion_Programada then
        begin
          TimerActualizacionPreciosProgramada.Interval:= 420000;  // establezco el timer cada 7 minutos
          //dtFecha_Proceso_ActuPreciosProgramada := IncDay(DateOf(Now),-1);

          try
            if Not Assigned(_Gestor_Articulos) then
              _Gestor_Articulos := TGestor_Articulos.Create(oParametro);
          except
            on E: Exception do
            begin
              _Grabar_LogErrores_Comunicacion(TTimer(Sender).Tag, 'No se puede crear _Gestor_Articulos. Error: '+E.Message);
            end;
          end;

        end;
      end;

      if Self.Tag=1 then   // Solo en el tag = 1 se habilita el timer. Osea en un solo servicio de tag 1
      begin
        if oParametro.Zafiro_BI_Activado then
        begin
          TimerActualizacionZafiroBI.Interval:= 180000;  // establezco el timer cada 3 minutos
          // TimerActualizacionZafiroBI.Interval:= 3000;  // establezco el timer cada 5 minutos

        end;
      end;

      if Self.Tag=1 then   // Solo en el tag = 1 se habilita el timer. Osea en un solo servicio de tag 1
      begin
        if oParametro.Costos_activa_recalculo then
        begin
          if oParametro.Costos_Cant_dias_recalculo>0 then
          begin
            TimerActualizacionPCpraPP.Interval:= 420000;  // establezco el timer cada 7 minutos

            try
              if Not Assigned(_Gestor_Articulos) then
                _Gestor_Articulos := TGestor_Articulos.Create(oParametro);
            except
              on E: Exception do
              begin
                _Grabar_LogErrores_Comunicacion(TTimer(Sender).Tag, 'No se puede crear _Gestor_Articulos. Error: '+E.Message);
              end;
            end;
          end;
        end;
      end;


      if Self.Tag=1 then   // Solo en el tag = 1 se habilita el timer. Osea en un solo servicio de tag 1
      begin
        if (oParametro.Fact_Online_FTP_Host <> '')
          and (oParametro.Fact_Online_FTP_User <> '')
            and (oParametro.Fact_Online_FTP_Pass <> '')
              and (oParametro.Fact_Online_FTP_Port > 0)
                and (oParametro.Fact_Online_FTP_URL_base <> '') then
        begin
          TimerSubidaAutoFTPCompPend.Interval:= 300000;  // establezco el timer cada 5 minutos
        end;
      end;


      bHaySucursales := False;
      if Self.Tag=1 then   // Solo en el tag = 1 se habilita el timer. Osea en un solo servicio de tag 1
      begin
        // Saco la Hora actual de la BD
        if qrySelect_Now.Active then
          qrySelect_Now.Close;
        qrySelect_Now.Open;
        if qrySelect_Now.RecordCount > 0 then
          dtFechaActual:= qrySelect_Now.FieldByName('Fecha_Hora').AsDateTime
        else
          dtFechaActual:= Now;
        qrySelect_Now.Close;
        //

        DecodeDateTime(dtFechaActual, wAnio, wMes, wDia, wHora, wMin, wSeg, wMil);

        if oParametro.Dia_Mes_Actualiza_MinMax = wDia then
        begin
          dsSucursales := nil;
          if Assigned(_Gestor_Sucursal) then
            dsSucursales := _Gestor_Sucursal._Get_Todas_Sucursales;

          if Assigned(dsSucursales) then
          begin
            if dsSucursales.Active then
              dsSucursales.Close;
            dsSucursales.Open;
            if dsSucursales.RecordCount > 0 then
            begin
              dsSucursales.First;
              while not(dsSucursales.Eof) do
              begin
                if dsSucursales.FieldByName('habilita_calc_minmax_autom').AsInteger > 0 then
                begin
                  bHaySucursales := True;
                  Break;
                end;
                dsSucursales.Next;
              end;
            end;
          end;

          if bHaySucursales then
          begin
            if Self.Tag=1 then   // Solo en el tag = 1 se habilita el timer. Osea en un solo servicio de tag 1
            begin
              TimerActualizacion_Min_Max_Stock.Interval:= 300000;  // establezco el timer cada 5 minutos

              try
                if Not Assigned(_Gestor_Articulos) then
                  _Gestor_Articulos := TGestor_Articulos.Create(oParametro);
              except
                on E: Exception do
                begin
                  _Grabar_LogErrores_Comunicacion(TTimer(Sender).Tag, 'No se puede crear _Gestor_Articulos. Error: '+E.Message);
                end;
              end;
            end;
          end;

        end;
      end;
    end;
  end;

  //este procedimiento crea una secuencia que activa o desactiva el temporizador

  // Habilito de acuerdo al Tag sel servicio
  _Hab_des_Timer(Self.Tag, True);

  TimerEtiquetas.Enabled:= False;
  if TimerEtiquetas.Interval > 0 then
  begin
    TimerEtiquetas.Enabled:= True;
    _Grabar_LogErrores_Comunicacion(-1, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Etiquetas Habilitado.'+Chr(13)+
      ' --> Carpeta Destino: ' + Trim(oParametro.Etiq_elec_carpeta_destino)+ Chr(13)+
      ' --> Periodicidad de la consulta: ' + IntToStr(oParametro.Etiq_elec_periodicidad_minutos)+ ' minutos'+ Chr(13)+
      ' --> Cantidad de días a contemplar cambios de precio: '+ IntToStr(oParametro.Etiq_elec_cant_dias_periodo));
    // Se hace la primera vez y no se tiene que esperar el tiempo del timer
    TimerEtiquetasTimer(Nil);
  end;

  TimerEsperaConsultaIQVIA.Enabled:= False;
  if TimerEsperaConsultaIQVIA.Interval > 0 then
  begin
    TimerEsperaConsultaIQVIA.Enabled:= True;
    _Grabar_LogErrores_Comunicacion(-2, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - IQVIA Habilitado.');
    //  TimerEsperaConsultaIQVIATimer(Nil);
  end;

  TimerExportarRecetasCloseup.Enabled:= False;
  if TimerExportarRecetasCloseup.Interval > 0 then
  begin
    TimerExportarRecetasCloseup.Enabled:= True;
    _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - CloseUp Habilitado.');
    // TimerExportarRecetasCloseupTimer(Nil);
  end;

  TimerExportacionDataview.Enabled:= False;
  if TimerExportacionDataview.Interval > 0 then
  begin
     TimerExportacionDataview.Enabled:= True;
     _Grabar_LogErrores_Comunicacion(-10, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - DataView Habilitado.');
     // TimerExportacionDataviewTimer(Nil);
  end;

  TimerActualizacionPreciosProgramada.Enabled:= False;
  if TimerActualizacionPreciosProgramada.Interval > 0 then
  begin
     TimerActualizacionPreciosProgramada.Enabled:= True;
     _Grabar_LogErrores_Comunicacion(-4, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Actualización de Precios Programada Habilitada.');
     //TimerActualizacionPreciosProgramadaTimer(Nil);
  end;

  TimerActualizacionZafiroBI.Enabled:= False;
  if TimerActualizacionZafiroBI.Interval > 0 then
  begin
     TimerActualizacionZafiroBI.Enabled:= True;
     _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Actualización Zafiro BI Habilitada.');
     // PV
     // TimerActualizacionZafiroBITimer(Nil);
  end;

  TimerActualizacionPCpraPP.Enabled:= False;
  if TimerActualizacionPCpraPP.Interval > 0 then
  begin
     TimerActualizacionPCpraPP.Enabled:= True;
     _Grabar_LogErrores_Comunicacion(-9, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Actualización de Precios de Costos Promedio Ponderado Habilitada.');
     //   TimerActualizacionPCpraPPTimer(Nil);
  end;


  TimerSubidaAutoFTPCompPend.Enabled:= False;
  if TimerSubidaAutoFTPCompPend.Interval > 0 then
  begin
    TimerSubidaAutoFTPCompPend.Enabled:= True;
    _Grabar_LogErrores_Comunicacion(-6, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Subida Automática de Comprobantes Pendientes en FTP Habilitada.');
    //    TimerSubidaAutoFTPCompPendTimer(Nil);
  end;

  TimerActualizacion_Min_Max_Stock.Enabled:= False;
  if TimerActualizacion_Min_Max_Stock.Interval > 0 then
  begin
    TimerActualizacion_Min_Max_Stock.Enabled:= True;
    _Grabar_LogErrores_Comunicacion(-7, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Actualización Áutomática de Min Max de Stock Habilitada.');
    //    TimerActualizacion_Min_Max_StockTimer(Nil);
  end;

  iSucursal_procesando1:= 0;
  iSucursal_procesando2:= 0;
  iSucursal_procesando3:= 0;
  iSucursal_procesando4:= 0;
  iSucursal_procesando5:= 0;
  iSucursal_procesando6:= 0;

  iSucursal_procesando7:= 0;
  iSucursal_procesando8:= 0;
  iSucursal_procesando9:= 0;
  iSucursal_procesando10:= 0;
  iSucursal_procesando11:= 0;
  iSucursal_procesando12:= 0;
  iSucursal_procesando13:= 0;

  iSucursal_procesando14:= 0;
  iSucursal_procesando15:= 0;
  iSucursal_procesando16:= 0;
  iSucursal_procesando17:= 0;

  iSucursal_procesando18:= 0;
  iSucursal_procesando19:= 0;
  iSucursal_procesando20:= 0;
  iSucursal_procesando21:= 0;
  iSucursal_procesando22:= 0;
  iSucursal_procesando23:= 0;
  iSucursal_procesando24:= 0;
  iSucursal_procesando25:= 0;
  iSucursal_procesando26:= 0;
  iSucursal_procesando27:= 0;
  iSucursal_procesando28:= 0;

  iSucursal_procesando100:= 0;

  bErrorComunicacion:= False;
  bErrorComunicacionConexionLocal:=False;
  while not Terminated do
  begin
    if Assigned(ServiceThread) then
    begin
      //_Grabar_LogErrores_Comunicacion(TTimer(Sender).Tag, 'Existe Hilo');
      ServiceThread.ProcessRequests(True);// wait for termination
    end
    else
    begin
      //_Grabar_LogErrores_Comunicacion(TTimer(Sender).Tag, 'No Existe Hilo');
      Break;
    end;
  end;

  _Hab_des_Timer(Self.Tag, False);

  TimerEtiquetas.Enabled:= False;

end;

procedure TVSComunicacionParaReplicacion01.ServiceAfterInstall(Sender: TService);
var
  Reg: TRegistry;
begin
  //Cargo la descripción al servicio para que se vea en el serviceManager y en el regedit
  Reg := TRegistry.Create(KEY_READ or KEY_WRITE);
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKey('\SYSTEM\CurrentControlSet\Services\' + Name, false) then
    begin
      Reg.WriteString('Description', 'V&S - Servicio de comunicación para replicación de datos entre servidores de BD remotos');
      Reg.CloseKey;
    end;
  finally
    Reg.Free;
  end;

end;

procedure TVSComunicacionParaReplicacion01.ServiceCreate(Sender: TObject);
var
  sNombre: String;
  iPosicion_caracter_inicial, iPosicion_caracter_final: Integer;
begin
  sNombre:= ParamStr(0);
  iPosicion_caracter_inicial:=PosEx('XEServicio_comunicacion', sNombre);
  if iPosicion_caracter_inicial> -1 then
  begin
    iPosicion_caracter_inicial := iPosicion_caracter_inicial + 23;
    sNombre:= copy(sNombre, iPosicion_caracter_inicial, length(sNombre)-iPosicion_caracter_inicial+1);
    iPosicion_caracter_final:=PosEx('.',sNombre);
    if iPosicion_caracter_final> -1 then
    begin
      sNombre:= Trim(copy(sNombre,0, (iPosicion_caracter_final-1)));
      try
        self.tag:= StrToInt(sNombre);
        self.DisplayName:=self.DisplayName + ' - Proceso '+sNombre;
      except
      end;
    end;
  end;
end;

procedure TVSComunicacionParaReplicacion01.Timer1Timer(Sender: TObject);
var
  sCadena_error: AnsiString;
  //bSalioErrorExepcionPrincipal: Boolean;
  iSucursal_procesando: Integer;
  bSigueEnLaMismaSucursal: Boolean;

begin
  //Deshabilito el timer para asegurar que termine el procedimiento cumpla o no el tiempo establecido

  _Hab_des_Timer(TTimer(Sender).Tag, False);

  try
    Conexion.Connected:= True;
    // Desactivo el log local para no registrar las transacciones de la comunicacion
    qryDesactivar_log_local.ExecSQL;
    bErrorComunicacionConexionLocal:= False;
//    CoolTrayIcon1.IconIndex:=1;
  except
    on E: Exception do
    begin
      //Utilizo esta variable bErrorComunicacionConexionLocal para grabar solo una vez el error de comunicacion si el problema persiste
      if bErrorComunicacionConexionLocal <> true then
      begin
        sCadena_error:= DateTimeToStr(Now)+'  No se puede conectar a la BD Local -  Error: '+ (E.Message);
        _Grabar_LogErrores_Comunicacion(TTimer(Sender).Tag, sCadena_error);
      end;

      bErrorComunicacionConexionLocal:= True;
//      CoolTrayIcon1.IconIndex:=0;
      _CargarItemsMenu;

      _Hab_des_Timer(TTimer(Sender).Tag, True);
      Exit;
    end;
  end;

  if not Assigned(oParametro) then
  begin
    //Conexion.Connected:= False;
//    CoolTrayIcon1.IconIndex:=0;
    _Hab_des_Timer(TTimer(Sender).Tag, True);
    Exit;
  end;

  if not Assigned(oSucursales) then
    oSucursales:= _Carga_Sucursales(TTimer(Sender).Tag, bErrorComunicacion);

  if Assigned(oSucursales) then
  begin
//    CoolTrayIcon1.IconIndex:=1;
  end
  else
  begin
    //Conexion.Connected:= False;
    bErrorComunicacion:= True;
//    CoolTrayIcon1.IconIndex:=0;
    _Hab_des_Timer(TTimer(Sender).Tag, True);
    Exit;
  end;

  //Actualizo el ultimo id_archivo_log_procesado y ultimo el pos_log_procesado es decir pongo el limite superior
  //que me indica la ultima transacción de la BD origen a la cual se tienen que empardar las demas BD
  if _Actualiza_Ultimo_Log_Procesado(TTimer(Sender).Tag) = False then
  begin
    //Conexion.Connected:= False;
//    CoolTrayIcon1.IconIndex:=0;
    _CargarItemsMenu;
    _Hab_des_Timer(TTimer(Sender).Tag, True);
    Exit;
  end;

  _CargarItemsMenu;

  if TTimer(Sender).Tag= 1 then
    iSucursal_procesando:= iSucursal_procesando1
  else if TTimer(Sender).Tag= 2 then
    iSucursal_procesando:= iSucursal_procesando2
  else if TTimer(Sender).Tag= 3 then
    iSucursal_procesando:= iSucursal_procesando3
  else if TTimer(Sender).Tag= 4 then
    iSucursal_procesando:= iSucursal_procesando4
  else if TTimer(Sender).Tag= 5 then
    iSucursal_procesando:= iSucursal_procesando5
  else if TTimer(Sender).Tag= 6 then
    iSucursal_procesando:= iSucursal_procesando6
  else if TTimer(Sender).Tag= 7 then
    iSucursal_procesando:= iSucursal_procesando7
  else if TTimer(Sender).Tag= 8 then
    iSucursal_procesando:= iSucursal_procesando8
  else if TTimer(Sender).Tag= 9 then
    iSucursal_procesando:= iSucursal_procesando9
  else if TTimer(Sender).Tag= 10 then
    iSucursal_procesando:= iSucursal_procesando10
  else if TTimer(Sender).Tag= 11 then
    iSucursal_procesando:= iSucursal_procesando11
  else if TTimer(Sender).Tag= 12 then
    iSucursal_procesando:= iSucursal_procesando12
  else if TTimer(Sender).Tag= 13 then
    iSucursal_procesando:= iSucursal_procesando13
  else if TTimer(Sender).Tag= 14 then
    iSucursal_procesando:= iSucursal_procesando14
  else if TTimer(Sender).Tag= 15 then
    iSucursal_procesando:= iSucursal_procesando15
  else if TTimer(Sender).Tag= 16 then
    iSucursal_procesando:= iSucursal_procesando16
  else if TTimer(Sender).Tag= 17 then
    iSucursal_procesando:= iSucursal_procesando17
  else if TTimer(Sender).Tag= 18 then
    iSucursal_procesando:= iSucursal_procesando18
  else if TTimer(Sender).Tag= 19 then
    iSucursal_procesando:= iSucursal_procesando19
  else if TTimer(Sender).Tag= 20 then
    iSucursal_procesando:= iSucursal_procesando20
  else if TTimer(Sender).Tag= 21 then
    iSucursal_procesando:= iSucursal_procesando21
  else if TTimer(Sender).Tag= 22 then
    iSucursal_procesando:= iSucursal_procesando22
  else if TTimer(Sender).Tag= 23 then
    iSucursal_procesando:= iSucursal_procesando23
  else if TTimer(Sender).Tag= 24 then
    iSucursal_procesando:= iSucursal_procesando24
  else if TTimer(Sender).Tag= 25 then
    iSucursal_procesando:= iSucursal_procesando25
  else if TTimer(Sender).Tag= 26 then
    iSucursal_procesando:= iSucursal_procesando26
  else if TTimer(Sender).Tag= 27 then
    iSucursal_procesando:= iSucursal_procesando27
  else if TTimer(Sender).Tag= 28 then
    iSucursal_procesando:= iSucursal_procesando28
  else if TTimer(Sender).Tag= 100 then
    iSucursal_procesando:= iSucursal_procesando100;
  // Registro de fecha hora para log de transmisión
//  if ((Self.Tag=1) or (Self.Tag=2) or (Self.Tag=3)) then   // Solo en el tag = 1 se ejecuta. Osea en un solo servicio de tag 1
//  begin
    sCadena_error:= _Modificar_Comu_Fecha_Hora(oParametro.Comu_Id_Sucursal);
    if Length(sCadena_error)>0 then
    begin
      sCadena_error:= DateTimeToStr(Now)+'  13---Sucursal:'+IntToStr(0)+' - '+'Sin Definir - '+ sCadena_error;
      _Grabar_LogErrores_Comunicacion(0, sCadena_error);
    end;

//  end;
  //



  while iSucursal_procesando <= (oSucursales.Count-1) do
  begin
  //for i:=0 to  do
  //begin
    if TSucursal(oSucursales.Items[iSucursal_procesando]).Exceptuado = False then
    begin
      if not (((TSucursal(oSucursales.Items[iSucursal_procesando]).Id_archivo_log_procesado = sUltimo_archivo_log_BDorigen) and (TSucursal(oSucursales.Items[iSucursal_procesando]).Pos_log_procesado >= iUltimo_Pos_log_BDorigen)) or (TSucursal(oSucursales.Items[iSucursal_procesando]).Id_archivo_log_procesado > sUltimo_archivo_log_BDorigen)) then
      begin
        if TSucursal(oSucursales.Items[iSucursal_procesando]).Id_ProcesoActualizacion=TTimer(Sender).Tag then
        begin
          try
            Conexion_Remota.ConnectionString:= TSucursal(oSucursales.Items[iSucursal_procesando]).Cadena_conexion;
            Conexion_Remota.Connected:= True;
            // Blanqueo el error de verificacion de Sentencias para que lo intente de nuevo
            TSucursal(oSucursales.Items[iSucursal_procesando]).ErrorVerificacionSentencias:=0;
            //

            //bSalioErrorExepcionPrincipal:= False;
            bSigueEnLaMismaSucursal:= True;
            while bSigueEnLaMismaSucursal=True do
            begin
              //_Grabar_LogErrores_Comunicacion(TSucursal(oSucursales.Items[iSucursal_procesando]).Id_ProcesoActualizacion, 'Inicia Proceso -> '+DateTimeToStr(Now)+'  Sucursal: '+IntToStr(TSucursal(oSucursales.Items[iSucursal_procesando]).Id_sucursal)+' - '+TSucursal(oSucursales.Items[iSucursal_procesando]).Des_sucursal);

              //if TSucursal(oSucursales.Items[iSucursal_procesando]).Id_ProcesoActualizacion=20 then // Ver de hacerlo con el timer 99
                  // Para recuperar datos de una tabla
              //  bSigueEnLaMismaSucursal:= _ActualizarBD(TSucursal(oSucursales.Items[iSucursal_procesando]), 'ubicaciones_por_articulo')
              //else
                bSigueEnLaMismaSucursal:= _ActualizarBD(TSucursal(oSucursales.Items[iSucursal_procesando]));

              //_Grabar_LogErrores_Comunicacion(TSucursal(oSucursales.Items[iSucursal_procesando]).Id_ProcesoActualizacion, 'Termino Proceso -> '+DateTimeToStr(Now)+'  Sucursal: '+IntToStr(TSucursal(oSucursales.Items[iSucursal_procesando]).Id_sucursal)+' - '+TSucursal(oSucursales.Items[iSucursal_procesando]).Des_sucursal);
            end;

          Except
            on E: Exception do
            begin
              sCadena_error:= DateTimeToStr(Now)+'  1---Sucursal:'+IntToStr(TSucursal(oSucursales.Items[iSucursal_procesando]).Id_sucursal)+' - '+TSucursal(oSucursales.Items[iSucursal_procesando]).Des_sucursal+' -  Error: '+ E.Message;
              //bSalioErrorExepcionPrincipal:= True;

              if TSucursal(oSucursales.Items[iSucursal_procesando]).GraboError= False then
                _Grabar_LogErrores_Comunicacion(TSucursal(oSucursales.Items[iSucursal_procesando]).Id_ProcesoActualizacion, sCadena_error);
              TSucursal(oSucursales.Items[iSucursal_procesando]).GraboError:= True;
              //Si una sucursal me da error en esta parte del codigo o en cualquier otro lado que quiera hacer una operación de BD
              //la exceptuo durante todo el servicio, y solo se la accedera de vuelta cuando el servicio sea reseteado o reiniciado.
              _CargarItemsMenu;

              //if RightStr(E.Message,26)='MySQL server has gone away' then
              //begin
                // Salgo del timer para que vuelva de nuevo pero en la sucursal siguiente
              //  Conexion_Remota.close;
              //  Conexion_Remota.Connected:= False;
              //  Conexion_Remota.ConnectionString:= '';
              //end;
              {
              Conexion_Remota.close;
              Conexion_Remota.Connected:= False;
              Conexion_Remota.ConnectionString:= '';
              iSucursal_procesando:= iSucursal_procesando + 1;
              if iSucursal_procesando > (oSucursales.Count-1) then
                iSucursal_procesando:=0;

              if TTimer(Sender).Tag= 1 then
                iSucursal_procesando1:= iSucursal_procesando
              else if TTimer(Sender).Tag= 2 then
                iSucursal_procesando2:= iSucursal_procesando
              else if TTimer(Sender).Tag= 3 then
                iSucursal_procesando3:= iSucursal_procesando
              else if TTimer(Sender).Tag= 4 then
                iSucursal_procesando4:= iSucursal_procesando
              else if TTimer(Sender).Tag= 5 then
                iSucursal_procesando5:= iSucursal_procesando
              else if TTimer(Sender).Tag= 6 then
                iSucursal_procesando6:= iSucursal_procesando
              else if TTimer(Sender).Tag= 7 then
                iSucursal_procesando7:= iSucursal_procesando
              else if TTimer(Sender).Tag= 8 then
                iSucursal_procesando8:= iSucursal_procesando
              else if TTimer(Sender).Tag= 9 then
                iSucursal_procesando9:= iSucursal_procesando
              else if TTimer(Sender).Tag= 10 then
                iSucursal_procesando10:= iSucursal_procesando
              else if TTimer(Sender).Tag= 11 then
                iSucursal_procesando11:= iSucursal_procesando
              else if TTimer(Sender).Tag= 12 then
                iSucursal_procesando12:= iSucursal_procesando;

              _Hab_des_Timer(TTimer(Sender).Tag, True);
              Exit;
              }
            end;
          end;
          Conexion_Remota.close;
          Conexion_Remota.Connected:= False;
          Conexion_Remota.ConnectionString:= '';
        end;
      end;
    end;    //if TSucursal(oSucursales.Items[iSucursal_procesando]).Exceptuado = False then

    // Siguiente Sucursal
    iSucursal_procesando:= iSucursal_procesando + 1;

  end;//end while iSucursal_procesando <= (oSucursales.Count-1) do
  iSucursal_procesando:= 0;
  if TTimer(Sender).Tag= 1 then
    iSucursal_procesando1:= iSucursal_procesando
  else if TTimer(Sender).Tag= 2 then
    iSucursal_procesando2:= iSucursal_procesando
  else if TTimer(Sender).Tag= 3 then
    iSucursal_procesando3:= iSucursal_procesando
  else if TTimer(Sender).Tag= 4 then
    iSucursal_procesando4:= iSucursal_procesando
  else if TTimer(Sender).Tag= 5 then
    iSucursal_procesando5:= iSucursal_procesando
  else if TTimer(Sender).Tag= 6 then
    iSucursal_procesando6:= iSucursal_procesando
  else if TTimer(Sender).Tag= 7 then
    iSucursal_procesando7:= iSucursal_procesando
  else if TTimer(Sender).Tag= 8 then
    iSucursal_procesando8:= iSucursal_procesando
  else if TTimer(Sender).Tag= 9 then
    iSucursal_procesando9:= iSucursal_procesando
  else if TTimer(Sender).Tag= 10 then
    iSucursal_procesando10:= iSucursal_procesando
  else if TTimer(Sender).Tag= 11 then
    iSucursal_procesando11:= iSucursal_procesando
  else if TTimer(Sender).Tag= 12 then
    iSucursal_procesando12:= iSucursal_procesando
  else if TTimer(Sender).Tag= 13 then
    iSucursal_procesando13:= iSucursal_procesando
  else if TTimer(Sender).Tag= 14 then
    iSucursal_procesando14:= iSucursal_procesando
  else if TTimer(Sender).Tag= 15 then
    iSucursal_procesando15:= iSucursal_procesando
  else if TTimer(Sender).Tag= 16 then
    iSucursal_procesando16:= iSucursal_procesando
  else if TTimer(Sender).Tag= 17 then
    iSucursal_procesando17:= iSucursal_procesando
  else if TTimer(Sender).Tag= 18 then
    iSucursal_procesando18:= iSucursal_procesando
  else if TTimer(Sender).Tag= 19 then
    iSucursal_procesando19:= iSucursal_procesando
  else if TTimer(Sender).Tag= 20 then
    iSucursal_procesando20:= iSucursal_procesando
  else if TTimer(Sender).Tag= 21 then
    iSucursal_procesando21:= iSucursal_procesando
  else if TTimer(Sender).Tag= 22 then
    iSucursal_procesando22:= iSucursal_procesando
  else if TTimer(Sender).Tag= 23 then
    iSucursal_procesando23:= iSucursal_procesando
  else if TTimer(Sender).Tag= 24 then
    iSucursal_procesando24:= iSucursal_procesando
  else if TTimer(Sender).Tag= 25 then
    iSucursal_procesando25:= iSucursal_procesando
  else if TTimer(Sender).Tag= 26 then
    iSucursal_procesando26:= iSucursal_procesando
  else if TTimer(Sender).Tag= 27 then
    iSucursal_procesando27:= iSucursal_procesando
  else if TTimer(Sender).Tag= 28 then
    iSucursal_procesando28:= iSucursal_procesando
  else if TTimer(Sender).Tag= 100 then
    iSucursal_procesando100:= iSucursal_procesando;

  _Hab_des_Timer(TTimer(Sender).Tag, True);
end;

procedure TVSComunicacionParaReplicacion01.TimerActualizacionPCpraPPTimer(
  Sender: TObject);
var
  oGestor_Lanzador_Zafiro: TGestor_Lanzador_Zafiro;
  dtFechaHoraActual: TDateTime;
  sRespuesta : String;
  oGestor_Parametros : TGestor_Parametros;
  dtFechaHoraUltimaActualizacion: TDateTime;
  dtFechaHoraProximaActualizacion: TDateTime;
  dtHoraProximaActualizacion: TDateTime;

  sHora_Actualizacion: String;
  dsArticulos_PPP : TDataSet;
begin

  // Deshabilito
  TimerActualizacionPCpraPP.Enabled:= False;
  //
  oGestor_Lanzador_Zafiro := TGestor_Lanzador_Zafiro.Create;
  dtFechaHoraActual:= oGestor_Lanzador_Zafiro._Fecha_Actual;

  dtFechaHoraUltimaActualizacion:= oParametro.Costos_Fecha_Hora_Ultima_Actualizacion;

  // Constante. No se configura. Ver si se debe configurar
  sHora_Actualizacion:= '00:10';
  //

//  dtFechaHoraUltimaActualizacion:= StrToDateTime('21/06/2023 00:15:00');
//  dtFechaHoraActual             := StrToDateTime('22/06/2023 00:16:00');

  dtHoraProximaActualizacion:= 0;
  if (dtFechaHoraActual - dtFechaHoraUltimaActualizacion)>1  then
  begin
    try
      dtHoraProximaActualizacion:= StrToTime(LeftStr(sHora_Actualizacion,5));
    except
      dtHoraProximaActualizacion:= 0;
    end;
    dtFechaHoraUltimaActualizacion:= DateOf(dtFechaHoraActual) - 1 + dtHoraProximaActualizacion;
  end;


  if dtFechaHoraUltimaActualizacion <> 0  then
  begin
    try
      dtHoraProximaActualizacion:= StrToTime(LeftStr(sHora_Actualizacion,5));    // A las 00:10 Hs
    except
      dtHoraProximaActualizacion:= 0;
    end;
    dtFechaHoraProximaActualizacion:= DateOf(dtFechaHoraUltimaActualizacion) + 1 + dtHoraProximaActualizacion;
  end;


  if dtFechaHoraActual > dtFechaHoraProximaActualizacion then
  begin
    dsArticulos_PPP := oGestor_Lanzador_Zafiro._Buscar_Precio_Compra_Promedio_Ponderado(DateOf(dtFechaHoraActual), oParametro.Costos_Cant_dias_recalculo);
    if Assigned(dsArticulos_PPP) then
    begin
      if dsArticulos_PPP.RecordCount > 0 then
      begin

        _Grabar_LogErrores_Comunicacion(-9, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Actualizando Precios de Compra Promedio Ponderado...');

        sRespuesta:= _Gestor_Articulos._Modificar_Articulos_Precio_Compra_Promedio_Ponderado(dsArticulos_PPP, 50);

        if LeftStr(sRespuesta, 5) <> 'Error' then
        begin
          _Grabar_LogErrores_Comunicacion(-9, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - '+'Proceso finalizado exitosamente - Se actualizaron los precios de Compra de '+sRespuesta+' Artículos.');
          // Actualizar Prec_Fecha_Hora_Ultima_Actualizacion
          oGestor_Parametros := TGestor_Parametros.Create;
          sRespuesta:= oGestor_Parametros._Modificar_Costos_Fecha_Hora_Ultima_Actualizacion;
          if Length(sRespuesta)>0 then
            _Grabar_LogErrores_Comunicacion(-9, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - '+sRespuesta);
          oParametro.Costos_Fecha_Hora_Ultima_Actualizacion := Now();
          FreeAndNil(oGestor_Parametros);
          //
        end;
      end;
    end;
  end;

  if Assigned(oGestor_Lanzador_Zafiro) then
    FreeAndNil(oGestor_Lanzador_Zafiro);

  // Habilito Timer
  TimerActualizacionPCpraPP.Enabled:= True;

end;

procedure TVSComunicacionParaReplicacion01.TimerActualizacionPreciosProgramadaTimer(
  Sender: TObject);
var
  dtFechaHoraActual: TDateTime;
  sRespuesta : String;
  oGestor_Parametros : TGestor_Parametros;
  dtFechaHoraUltimaActualizacion: TDateTime;
  dtFechaHoraProximaActualizacion: TDateTime;
  dtHoraProximaActualizacion: TDateTime;
begin

  // Deshabilito
  TimerActualizacionPreciosProgramada.Enabled:= False;
  //

  // Saco la Hora actual de la BD
  if qrySelect_Now.Active then
    qrySelect_Now.Close;
  qrySelect_Now.Open;
  if qrySelect_Now.RecordCount > 0 then
    dtFechaHoraActual:= qrySelect_Now.FieldByName('Fecha_Hora').AsDateTime
  else
    dtFechaHoraActual:= Now;
  qrySelect_Now.Close;
  //


  dtFechaHoraUltimaActualizacion:= oParametro.Prec_Fecha_Hora_Ultima_Actualizacion;

//  dtFechaHoraUltimaActualizacion:= StrToDateTime('26/04/2023 02:00:00');
//  dtFechaHoraActual             := StrToDateTime('27/04/2023 03:01:00');

  if (dtFechaHoraActual - dtFechaHoraUltimaActualizacion)>1  then
  begin
    try
      dtHoraProximaActualizacion:= StrToTime(LeftStr(oParametro.Prec_Hora_Actualizacion,5));
    except
      dtHoraProximaActualizacion:= 0;
    end;
    dtFechaHoraUltimaActualizacion:= DateOf(dtFechaHoraActual) - 1 + dtHoraProximaActualizacion;
  end;


  if dtFechaHoraUltimaActualizacion <> 0  then
  begin
    try
      dtHoraProximaActualizacion:= StrToTime(LeftStr(oParametro.Prec_Hora_Actualizacion,5));
    except
      dtHoraProximaActualizacion:= 0;
    end;
    dtFechaHoraProximaActualizacion:= DateOf(dtFechaHoraUltimaActualizacion) + 1 + dtHoraProximaActualizacion;
  end;

  _Grabar_LogErrores_Comunicacion(-4, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Verificando horario... Fecha hora actual: '+FormatDateTime('yyyy-mm-dd hh:mm:ss', dtFechaHoraActual)+' - Fecha hora próxima actualización: '+FormatDateTime('yyyy-mm-dd hh:mm:ss', dtFechaHoraProximaActualizacion));

  if dtFechaHoraActual > dtFechaHoraProximaActualizacion then
  begin
    _Grabar_LogErrores_Comunicacion(-4, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Actualizando Precios Programados...');
    sRespuesta:= _Actualizar_Precios_Programados;
    if Length(sRespuesta)=0 then
    begin
      // Actualizar Prec_Fecha_Hora_Ultima_Actualizacion
      oGestor_Parametros := TGestor_Parametros.Create;
      oGestor_Parametros._Modificar_Prec_Fecha_Hora_Ultima_Actualizacion;
      oParametro.Prec_Fecha_Hora_Ultima_Actualizacion := Now();
      FreeAndNil(oGestor_Parametros);
      //
    end;
  end;


  // Habilito
  TimerActualizacionPreciosProgramada.Enabled:= True;
end;

procedure TVSComunicacionParaReplicacion01.TimerActualizacionZafiroBITimer(
  Sender: TObject);
var
  oGestor_Lanzador_Zafiro: TGestor_Lanzador_Zafiro;
  dtFechaHoraActual: TDateTime;
  sRespuesta : String;
  oGestor_Parametros : TGestor_Parametros;
  dtFechaHoraUltimaActualizacion: TDateTime;
  dtFechaHoraProximaActualizacion: TDateTime;
  dtHoraProximaActualizacion: TDateTime;
  oParametro_Zafiro_BI: TParametro_Zafiro_BI;

  sHora_Actualizacion: String;
begin

  // Deshabilito
  TimerActualizacionZafiroBI.Enabled:= False;
  //
  oGestor_Lanzador_Zafiro := TGestor_Lanzador_Zafiro.Create;
  dtFechaHoraActual:= oGestor_Lanzador_Zafiro._Fecha_Actual;

  dtFechaHoraUltimaActualizacion:= oParametro.Zafiro_BI_Fecha_Hora_Ultima_Actualizacion;

  // Constante. No se configura. Ver si se debe configurar
  sHora_Actualizacion:= '00:15';
  //

//  dtFechaHoraUltimaActualizacion:= StrToDateTime('21/06/2023 00:15:00');
//  dtFechaHoraActual             := StrToDateTime('22/06/2023 00:16:00');

  if (dtFechaHoraActual - dtFechaHoraUltimaActualizacion)>1  then
  begin
    try
      dtHoraProximaActualizacion:= StrToTime(LeftStr(sHora_Actualizacion,5));
    except
      dtHoraProximaActualizacion:= 0;
    end;
    dtFechaHoraUltimaActualizacion:= DateOf(dtFechaHoraActual) - 1 + dtHoraProximaActualizacion;
  end;


  if dtFechaHoraUltimaActualizacion <> 0  then
  begin
    try
      dtHoraProximaActualizacion:= StrToTime(LeftStr(sHora_Actualizacion,5));    // A las 00:15 Hs
    except
      dtHoraProximaActualizacion:= 0;
    end;
    dtFechaHoraProximaActualizacion:= DateOf(dtFechaHoraUltimaActualizacion) + 1 + dtHoraProximaActualizacion;
  end;


  if dtFechaHoraActual > dtFechaHoraProximaActualizacion then
  begin

    oParametro_Zafiro_BI:= oGestor_Lanzador_Zafiro._Buscar_Parametro_Zafiro_BI;

    if Assigned(oParametro_Zafiro_BI) then
    begin
      if oParametro_Zafiro_BI.Vs_Id_Cliente>0 then
      begin
        if Length(Trim(oParametro_Zafiro_BI.Cadena_Conexion_BD))>0 then
        begin
          try
            Conexion_Remota.ConnectionString:= oParametro_Zafiro_BI.Cadena_Conexion_BD;
            Conexion_Remota.Connected:= True;
            // Aqui
            sRespuesta:= _Actualizar_Zafiro_BI(oParametro_Zafiro_BI, DateOf(dtFechaHoraActual)-1);
            if Length(sRespuesta)>0 then
            begin
              // No Hago nada... Ya fue guardando en el log
              {if sRespuesta='Reiniciar' then
              begin
                //
                TimerActualizacionZafiroBI.Interval:= 500;  // establezco el timer en medio segundo para que reinicie inmediatamente
                //
                _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Proceso de actualizacion BD Zafiro BI reiniciado');
              end;}
            end
            else
            begin
              // Actualizar Zafiro_BI_Fecha_Hora_Ultima_Actualizacion ( z_bi_fua - Esta encriptada )

              // Ver de descomentar

              oGestor_Parametros := TGestor_Parametros.Create;
              sRespuesta:= oGestor_Parametros._Modificar_ZafiroBI_Fecha_Hora_Ultima_Actualizacion;
              oParametro.Zafiro_BI_Fecha_Hora_Ultima_Actualizacion := Now();
              FreeAndNil(oGestor_Parametros);

              if Length(sRespuesta)>0 then
                _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - '+sRespuesta);

              //TimerActualizacionZafiroBI.Interval:= 180000;  // establezco el timer cada 3 minutos
              //
            end;
            //
          except
            on E: Exception do
            begin
              _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - No se pudo establecer conexión con la BD Zafiro BI. (1) -  Error: '+ E.Message);
            end;
          end;
          Conexion_Remota.close;
          Conexion_Remota.Connected:= False;
          Conexion_Remota.ConnectionString:= '';
        end
        else
        begin
          _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - No se pudo obtener la cadena de conexión para Zafiro BI. (2)');
        end;
      end
      else
      begin
        _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - No se pudo obtener el cliente Zafiro para el BI. (3)');
      end;
      FreeAndNil(oParametro_Zafiro_BI);
    end
    else
    begin
      _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - No se pudo obtener los parametros del BI. (4)');
    end;
  end;

  if Assigned(oGestor_Lanzador_Zafiro) then
    FreeAndNil(oGestor_Lanzador_Zafiro);

  // Habilito
  TimerActualizacionZafiroBI.Enabled:= True;

end;

procedure TVSComunicacionParaReplicacion01.TimerActualizacion_Min_Max_StockTimer(
  Sender: TObject);
var
  oGestor_Parametros : TGestor_Parametros;
  dtFechaHoraActual: TDateTime;
  sRespuesta : String;
  dtFechaHoraUltimaActualizacion: TDateTime;
  dtFechaHoraProximaActualizacion: TDateTime;
  dtHoraProximaActualizacion: TDateTime;
  sHora_Actualizacion: String;
  dsEmpresas: TDataSet;
begin
  //Deshabilito
  TimerActualizacion_Min_Max_Stock.Enabled := False;

  // Saco la Hora actual de la BD
  if qrySelect_Now.Active then
    qrySelect_Now.Close;
  qrySelect_Now.Open;
  if qrySelect_Now.RecordCount > 0 then
    dtFechaHoraActual:= qrySelect_Now.FieldByName('Fecha_Hora').AsDateTime
  else
    dtFechaHoraActual:= Now;
  qrySelect_Now.Close;
  //

  oGestor_Parametros := TGestor_Parametros.Create;
  dtFechaHoraUltimaActualizacion:= oGestor_Parametros._Get_FechaHora_Ultima_Actu_Min_Max_Stock;

  // Constante. No se configura. Ver si se debe configurar
  sHora_Actualizacion:= '00:30';

  if (dtFechaHoraActual - dtFechaHoraUltimaActualizacion)>1  then
  begin
    try
      dtHoraProximaActualizacion:= StrToTime(LeftStr(sHora_Actualizacion,5));
    except
      dtHoraProximaActualizacion:= 0;
    end;
    dtFechaHoraUltimaActualizacion:= DateOf(dtFechaHoraActual) - 1 + dtHoraProximaActualizacion;
  end;

  if dtFechaHoraUltimaActualizacion <> 0  then
  begin
    try
      dtHoraProximaActualizacion:= StrToTime(LeftStr(sHora_Actualizacion,5));    // A las 00:30 Hs
    except
      dtHoraProximaActualizacion:= 0;
    end;
    dtFechaHoraProximaActualizacion:= DateOf(dtFechaHoraUltimaActualizacion) + 1 + dtHoraProximaActualizacion;
  end;

  if dtFechaHoraActual > dtFechaHoraProximaActualizacion then
  begin

    dsEmpresas := nil;
    if Assigned(_Gestor_Empresa) then
      dsEmpresas := _Gestor_Empresa._Get_Empresas;

    if Assigned(dsEmpresas) then
    begin
      if dsEmpresas.RecordCount > 0 then
      begin
        dsEmpresas.First;
        while not(dsEmpresas.Eof) do
        begin
          _Actualizar_Min_Max_Stock(dsEmpresas.FieldByName('id_empresa').AsInteger);
          dsEmpresas.Next;
        end;
      end;
      dsEmpresas.Close;
      FreeAndNil(dsEmpresas);
    end;


    sRespuesta:= oGestor_Parametros._Modificar_Ultima_Fecha_Hora_Actu_Min_Max_Stock;
    if Length(sRespuesta)>0 then
      _Grabar_LogErrores_Comunicacion(-7, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - '+sRespuesta);
  end;


  if Assigned(oGestor_Parametros) then
    FreeAndNil(oGestor_Parametros);

  //Habilito nuevamente
  TimerActualizacion_Min_Max_Stock.Enabled := True;
end;

procedure TVSComunicacionParaReplicacion01.TimerEtiquetasTimer(Sender: TObject);
var
  _InterfaceArchivos : TInterfaceArchivos;
  _Parametro_Formato : TParametro_Fotmato;

  dDiaHasta: TDate;
  dDiaDesde: TDate;
  sNombreArchivo: String;
begin
  try
    sNombreArchivo:= 'EtiquetasElectronicas_'+FormatDateTime('yyyymmddhhmmss', Now)+'.txt';
    //_Grabar_LogErrores_Comunicacion(0, 'Se inicia exportacion de registros en el archivo '+oParametro.Etiq_elec_carpeta_destino+'\'+sNombreArchivo );

    with qrySelectAuditoria_Articulo_Resumido do
    begin
      if Active then
        Close;

      dDiaHasta:= Now;
      dDiaDesde:= dDiaHasta - oParametro.etiq_elec_cant_dias_periodo;

      // ** paso parametro consulta ** //
      //Parameters.ParamByName('pId_empresa').Value:= oParametros.Id_Empresa;
      //Parameters.ParamByName('pFechaHoraDesde1').Value:=FormatDateTime('yyyy/mm/dd', dDiaDesde);
      //Parameters.ParamByName('pFechaHoraHasta1').Value:=FormatDateTime('yyyy/mm/dd', dDiaHasta+1);

      Parameters.ParamByName('pFechaHoraDesde2').Value:=FormatDateTime('yyyy/mm/dd', dDiaDesde);
      Parameters.ParamByName('pFechaHoraHasta2').Value:=FormatDateTime('yyyy/mm/dd', dDiaHasta+1);

      Parameters.ParamByName('pId_empresa6').Value  := oParametro.Id_Empresa;

      SQL[6]  := '';
      SQL[19] := '';
      SQL[20] := '';

      SQL[21] := '';
      SQL[23] := '';

      SQL[SQL.Count-1]:='ORDER BY id_articulo';
      Open;
    end; // with qrySelectAuditoria_Articulo_Resumido

    if qrySelectAuditoria_Articulo_Resumido.RecordCount > 0  then
    begin
      _Parametro_Formato := TParametro_Fotmato.Create;
      if Length(Trim(_Parametro_Formato.Formato_exportacion_etiqueta_electronica))>0 then
      begin
        _InterfaceArchivos := TInterfaceArchivos.Create(_Parametro_Formato.Formato_exportacion_etiqueta_electronica);
        if Length(Trim(_Parametro_Formato.Formato_exportacion_etiqueta_electronica))>0 then
        begin
          if Assigned(_InterfaceArchivos) then
          begin

            if _InterfaceArchivos.Exportar(Nil,qrySelectAuditoria_Articulo_Resumido,Nil, oParametro.Etiq_elec_carpeta_destino+'\', sNombreArchivo, True) then
            begin
              _Grabar_LogErrores_Comunicacion(-1, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Se exportaron '+IntToStr(qrySelectAuditoria_Articulo_Resumido.RecordCount)+' registros en el archivo '+oParametro.Etiq_elec_carpeta_destino+'\'+sNombreArchivo );
            end
            else
            begin
              _Grabar_LogErrores_Comunicacion(-1, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+' - No se pudo generar el archivo '+oParametro.Etiq_elec_carpeta_destino+'\'+sNombreArchivo);
            end;
          end;
        end;
        FreeAndNil(_InterfaceArchivos);
      end
      else
      begin
        //Mensaje_Error('No está configurado el formato de exportación de Cenefas Electrónicas. No se puede generar el archivo');
        _Grabar_LogErrores_Comunicacion(-1, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+' - No está configurado el formato de exportación de Etiquetas Electrónicas. No se puede generar el archivo');
      end;
      FreeAndNil(_Parametro_Formato);
    end
    else
    begin
      _Grabar_LogErrores_Comunicacion(-1, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - No existen datos para exportar');
    end;

    qrySelectAuditoria_Articulo_Resumido.Close;
  except
  end;

end;

procedure TVSComunicacionParaReplicacion01.TimerExportacionDataviewTimer(
  Sender: TObject);
begin
  _Exportacion_DataView;
end;

procedure TVSComunicacionParaReplicacion01.TimerExportarRecetasCloseupTimer(
  Sender: TObject);
begin
  _Exportar_Datos_Receta;
end;

procedure TVSComunicacionParaReplicacion01.TimerSubidaAutoFTPCompPendTimer(
  Sender: TObject);
var
  oGestor_Parametros : TGestor_Parametros;
  dtFechaHoraActual: TDateTime;
  sRespuesta : String;
  dtFechaHoraUltimaSubida: TDateTime;
  dtFechaHoraProximaSubida: TDateTime;
  dtHoraProximaSubida: TDateTime;
  sHora_Subida: String;
begin
  //Deshabilito
  TimerSubidaAutoFTPCompPend.Enabled := False;

  // Saco la Hora actual de la BD
  if qrySelect_Now.Active then
    qrySelect_Now.Close;
  qrySelect_Now.Open;
  if qrySelect_Now.RecordCount > 0 then
    dtFechaHoraActual:= qrySelect_Now.FieldByName('Fecha_Hora').AsDateTime
  else
    dtFechaHoraActual:= Now;
  qrySelect_Now.Close;
  //

  oGestor_Parametros := TGestor_Parametros.Create;
  dtFechaHoraUltimaSubida:= oGestor_Parametros._Get_FechaHora_ultima_subida_FTP;

  // Constante. No se configura. Ver si se debe configurar
  sHora_Subida:= '00:20';

  if (dtFechaHoraActual - dtFechaHoraUltimaSubida)>1  then
  begin
    try
      dtHoraProximaSubida:= StrToTime(LeftStr(sHora_Subida,5));
    except
      dtHoraProximaSubida:= 0;
    end;
    dtFechaHoraUltimaSubida:= DateOf(dtFechaHoraActual) - 1 + dtHoraProximaSubida;
  end;

  if dtFechaHoraUltimaSubida <> 0  then
  begin
    try
      dtHoraProximaSubida:= StrToTime(LeftStr(sHora_Subida,5));    // A las 00:15 Hs
    except
      dtHoraProximaSubida:= 0;
    end;
    dtFechaHoraProximaSubida:= DateOf(dtFechaHoraUltimaSubida) + 1 + dtHoraProximaSubida;
  end;

  if dtFechaHoraActual > dtFechaHoraProximaSubida then
  begin
    _Subir_Comp_Pend_Servidor_FTP;

    sRespuesta:= oGestor_Parametros._Modificar_Ultima_Fecha_Hora_Subida_Comp_FTP;
    if Length(sRespuesta)>0 then
      _Grabar_LogErrores_Comunicacion(-6, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - '+sRespuesta);
  end;


  if Assigned(oGestor_Parametros) then
    FreeAndNil(oGestor_Parametros);

  //Habilito nuevamente
  TimerSubidaAutoFTPCompPend.Enabled := True;
end;

procedure TVSComunicacionParaReplicacion01.TimerEsperaConsultaIQVIATimer(
  Sender: TObject);
var
  dtFecha       : TDate;
  dtFecha_Inicio: TDate;
  oGestor_Lanzador_Zafiro: TGestor_Lanzador_Zafiro;

  dtFecha_Inicio_IQVIa: TDate;
  oFormato_Fecha          : TFormatSettings;

  slConfig_iqvia_externo: TStringList;
  iLinea: SmallInt;
  iCaracter, iNroCaracterSeparador: Integer;
  sCadena, sRestoCadena, sFraccionCadena: string;
  sSeparadorCampos: String;

  iId_Empresa_Archivo: Integer;
  iId_Sucursal_Archivo: Integer;
  sDes_Sucursal_Archivo: String;
  sUbicacion_archivos_origen: String;
  sUbicacion_archivos_procesados: String;

  sRespuesta: String;

begin
  // extraerlo de parametros
  oFormato_Fecha.DateSeparator := #47;
  oFormato_Fecha.ShortDateFormat:='yyyy'+oFormato_Fecha.DateSeparator+'mm'+oFormato_Fecha.DateSeparator + 'dd';
  dtFecha_Inicio_IQVIa:= StrToDate('2021/06/01', oFormato_Fecha);
  // extraerlo de parametros


  // Deshabilito
  TimerEsperaConsultaIQVIA.Enabled:= False;
  //

  //tmrEsperaConsultaIQVIA.Interval:= 180000;  // establezco el timer cada 3 minutos (180000 milisegundos)

  oGestor_Lanzador_Zafiro := TGestor_Lanzador_Zafiro.Create;

  dtFecha:= DateOf(oGestor_Lanzador_Zafiro._Fecha_Actual);
  dtFecha_Inicio := IncDay(dtFecha,(-1)* iCantidad_Dias_Consulta_Iqvia);

  if dtFecha_Inicio<dtFecha_Inicio_IQVIa then
    dtFecha_Inicio:=dtFecha_Inicio_IQVIa;


  if dtFecha_Proceso_IQVIa <> dtFecha then
  begin
    if FileExists('c:\vs-comunicacion\config_iqvia_externo.txt') then
    begin
      slConfig_iqvia_externo:= TStringList.Create;
      slConfig_iqvia_externo.LoadFromFile('c:\vs-comunicacion\config_iqvia_externo.txt');
      if slConfig_iqvia_externo.Count>1 then
      begin
        for iLinea := 1 to slConfig_iqvia_externo.Count-1 do
        begin
          //iNroCaracterSeparador :=  Pos(#9, slConfig_iqvia_externo[iLinea]);
          sCadena:= slConfig_iqvia_externo[iLinea];
          iCaracter:=1;
          sSeparadorCampos:= #9;

          // id empresa
          sRestoCadena:= copy(sCadena, iCaracter, length(sCadena)-iCaracter+1);
          iNroCaracterSeparador :=  Pos(sSeparadorCampos, sRestoCadena);
          if iNroCaracterSeparador=0 then
            iNroCaracterSeparador:= length(sCadena);
          sFraccionCadena:= Copy(sRestoCadena, 1, iNroCaracterSeparador-1);
          iId_Sucursal_Archivo := 0;
          if Length(sFraccionCadena)>0 then
          begin
            try
              iId_Empresa_Archivo  := StrToInt(sFraccionCadena);
            except
              iId_Empresa_Archivo  := 0;
            end;
          end;
          iCaracter:= iCaracter + Length(sFraccionCadena)+1;
          //

          // id sucursal
          sRestoCadena:= copy(sCadena, iCaracter, length(sCadena)-iCaracter+1);
          iNroCaracterSeparador :=  Pos(sSeparadorCampos, sRestoCadena);
          if iNroCaracterSeparador=0 then
            iNroCaracterSeparador:= length(sCadena);
          sFraccionCadena:= Copy(sRestoCadena, 1, iNroCaracterSeparador-1);
          iId_Sucursal_Archivo := 0;
          if Length(sFraccionCadena)>0 then
          begin
            try
              iId_Sucursal_Archivo  := StrToInt(sFraccionCadena);
            except
              iId_Sucursal_Archivo  := 0;
            end;
          end;
          iCaracter:= iCaracter + Length(sFraccionCadena)+1;
          //

          // Des sucursal
          sRestoCadena:= copy(sCadena, iCaracter, length(sCadena)-iCaracter+1);
          iNroCaracterSeparador :=  Pos(sSeparadorCampos, sRestoCadena);
          if iNroCaracterSeparador=0 then
            iNroCaracterSeparador:= length(sCadena);
          sFraccionCadena:= Copy(sRestoCadena, 1, iNroCaracterSeparador-1);
          sDes_Sucursal_Archivo := '';
          if Length(sFraccionCadena)>0 then
            sDes_Sucursal_Archivo := sFraccionCadena;
          iCaracter:= iCaracter + Length(sFraccionCadena)+1;
          //

          // Ubicacion archivos origen
          sRestoCadena:= copy(sCadena, iCaracter, length(sCadena)-iCaracter+1);
          iNroCaracterSeparador :=  Pos(sSeparadorCampos, sRestoCadena);
          if iNroCaracterSeparador=0 then
            iNroCaracterSeparador:= length(sCadena);
          sFraccionCadena:= Copy(sRestoCadena, 1, iNroCaracterSeparador-1);
          sUbicacion_archivos_origen := '';
          if Length(sFraccionCadena)>0 then
            sUbicacion_archivos_origen := sFraccionCadena;
          iCaracter:= iCaracter + Length(sFraccionCadena)+1;
          //

          // Ubicacion archivos procesados
          sRestoCadena:= copy(sCadena, iCaracter, length(sCadena)-iCaracter+1);
          iNroCaracterSeparador :=  Pos(sSeparadorCampos, sRestoCadena);
          if iNroCaracterSeparador=0 then
            iNroCaracterSeparador:= length(sCadena);
          sFraccionCadena:= Copy(sRestoCadena, 1, iNroCaracterSeparador-1);
          sUbicacion_archivos_procesados := '';
          if Length(sFraccionCadena)>0 then
            sUbicacion_archivos_procesados := sFraccionCadena;
          iCaracter:= iCaracter + Length(sFraccionCadena)+1;
          //
          if ((iId_Empresa_Archivo>0) and (iId_Sucursal_Archivo>0) and (Length(sUbicacion_archivos_origen)>0) and (Length(sUbicacion_archivos_procesados)>0)) then
          begin
            _Informar_IQVIA_externo(iId_Empresa_Archivo, iId_Sucursal_Archivo, sUbicacion_archivos_origen, sUbicacion_archivos_procesados);
          end;
        end;
      end;
      FreeAndNil(slConfig_iqvia_externo);


    end
    else
    begin
      //_Consultar_Cupones_No_Informados_IQVIA('WS', dtFecha_Inicio, IncDay(dtFecha,-1));
    end;
    // Lo hago siempre
    _Consultar_Cupones_No_Informados_IQVIA('WS', dtFecha_Inicio, IncDay(dtFecha,-1));
    //
    // Para que no lo haga hasta el proximo día
    dtFecha_Proceso_IQVIa := DateOf(oGestor_Lanzador_Zafiro._Fecha_Actual);
    //
  end;

  if Assigned(oGestor_Lanzador_Zafiro) then
    FreeAndNil(oGestor_Lanzador_Zafiro);

  // Habilito
  TimerEsperaConsultaIQVIA.Enabled:= True;
end;

function TVSComunicacionParaReplicacion01._ActualizarBD(pSucursal: TSucursal; pTabla_Exportar: String = ''): Boolean;    // Devuelve si debe seguir procesando la misma sucursal
var
  sCadena, sNombreTabla, sNombreBD : String;
  sCadena_error : AnsiString;
  iNro_Extension, iPosicion_caracter_inicial, iPosicion_caracter_final, iPosicion_nombreBD, i, j: Integer;
  iNro_Extension_Local, iNro_Extension_Destino: Integer;

  bPrimer_archivo_log, bBD_erronea, bFin_archivo: Boolean;
  iEnd_log_pos: Longint;
  iEnd_log_pos_real: Longint;

  //sUltimo_archivo_log_BDestino: String;
  //iUltimo_Pos_log_BDestino    : Longint;
  iUltimo_Pos_log_BDestinoReal: Longint;

  iPosIni                     : Longint; // usado para grabar en el log la posición de inicio de las cadenas SQL que generaron el error

  sTextoEnriquecido: WideString;
  sSqlOriginal: String;
  sWhere: String;
  iSentencia_Dep_Texto_Enriquecido: Integer;
  ArrSentencias_Dep_Texto_Enriquecido: TArrInteger;
  bHayRegistros: Boolean;

  sId_comprobante: String;
  sId_empresa: String;
  sId_proveedor: String;
  sId_Persona: String;
  sId_Sucursal: String;
  sNro_caja: String;
  sPeriodo_caja: String;
  sSentencia: String;

  slSentenciasManuales: TStringList;

  bComprobanteTemporal: Boolean;
  bExcluyeCadena: Boolean;
  bEsAnnotate_rows: Boolean;

  bEsSucursales_comunicacionFecha_Hora: Boolean;

  bHayInsert: Boolean;
  iCantidad: Integer;

begin
  Result:= False;
  bPrimer_archivo_log:= False;
  qrySelect_Binlog_events.Close;
  qrySelect_Binlog_events.SQL.Clear;
  qrySentencias_comunicacion.SQL.Clear;
  bBD_erronea:= False;
  slSentencias_comunicacion:= TStringList.Create;
  iPosIni:= -1;

  If Length(Trim(pSucursal.Id_archivo_log_procesado))=0 then
  begin
    iNro_Extension:= 1;
    //'mysql-bin-vs' que esta en parametros pParametros.comu_nombre_file_binlog
    pSucursal.Id_archivo_log_procesado:= oParametro.Comu_nombre_file_binlog + '.'+ RightStr('00000000000000000'+IntToStr(iNro_Extension), oParametro.Comu_largo_extension_file_binlog);
    bPrimer_archivo_log:= True;
  end;

  // De bucle anterior debo verificar si chequeo la sentencia... Si no va con error.
  // Si pudo verificar las sentencias en la BD destino: grabo en el log local.
  if pSucursal.ErrorVerificacionSentencias = 0 then
  begin
    if Length(pTabla_Exportar)=0 then
    begin
      // Chequeo si el Log Procesado En Destino Es Mayor Que el Local
      // Si es asi, tomo el log del destino.
      // Quiere decir que no se actualizo bien cuando debe actualizar el log local en sucursales (corte abrupto de comunicacion)
      with qrySelectLogProcesadoOrigenEnDestino do
      begin
        Parameters.ParamByName('pId_sucursal').Value := oParametro.Comu_Id_Sucursal;
        try
          Open;
          if RecordCount > 0 then
          begin
            if not ((Trim(pSucursal.Id_archivo_log_procesado)=Trim(FieldByName('id_archivo_log_procesado_origen').AsString)) and (pSucursal.Pos_log_procesado=FieldByName('pos_log_procesado_origen').AsInteger)) then
            begin
              iNro_Extension_Local:= 0;
              iNro_Extension_Destino:= 0;
              try
                iNro_Extension_Local:= StrToInt(RightStr(pSucursal.Id_archivo_log_procesado, oParametro.Comu_largo_extension_file_binlog));
                iNro_Extension_Destino:= StrToInt(RightStr(Trim(FieldByName('id_archivo_log_procesado_origen').AsString), oParametro.Comu_largo_extension_file_binlog));
              except
                iNro_Extension_Local:= 0;
                iNro_Extension_Destino:= 0;
              end;

              if iNro_Extension_Destino > iNro_Extension_Local then
              begin
                pSucursal.Id_archivo_log_procesado:= FieldByName('id_archivo_log_procesado_origen').AsString;
                pSucursal.Pos_log_procesado       :=FieldByName('pos_log_procesado_origen').AsInteger;
                bPrimer_archivo_log:= False;
              end
              else
              begin
                if iNro_Extension_Destino = iNro_Extension_Local then
                begin
                  if FieldByName('pos_log_procesado_origen').AsInteger > pSucursal.Pos_log_procesado then
                  begin
                    pSucursal.Id_archivo_log_procesado:= FieldByName('id_archivo_log_procesado_origen').AsString;
                    pSucursal.Pos_log_procesado       :=FieldByName('pos_log_procesado_origen').AsInteger;
                    bPrimer_archivo_log:= False;
                  end;
                end;
              end;
            end;
          end;
          close;
        except
          on E: Exception do
          begin
            sCadena_error:= DateTimeToStr(Now)+'  9---Sucursal:'+IntToStr(pSucursal.Id_sucursal)+' - '+pSucursal.Des_sucursal+' -  Error: No se pudo obtener el ultimo Pos grabado en el destino. '+ (E.Message);

            //pSucursal.Exceptuado:= True;
            if pSucursal.GraboError= False then
              _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
            pSucursal.GraboError:= True;
            _CargarItemsMenu;
            if Assigned(slSentencias_comunicacion) then
              slSentencias_comunicacion.Free;
            Exit;
          end;
        end;

      end;//end with qrySelectLogProcesadoOrigenEnDestino do
    end;
    //
    //
  end
  else
  begin
    // Nada. Toma el actual
    //pSucursal.ErrorVerificacionSentencias := 0;
  end;  //if pSucursal.ErrorVerificacionSentencias = 0 then

  //Comienzo a leer el archivo desde el principio si bPrimer_archivo_log= true o desde la ultima posicióno que guarde en el pos_log_procesado
  if (bPrimer_archivo_log = True) or (pSucursal.Pos_log_procesado= 0) then
  begin
    with qrySelect_Binlog_events do
    begin
      //sql.Add('SHOW BINLOG EVENTS IN '+ QuotedStr(pSucursal.Id_archivo_log_procesado) +' LIMIT 400000');
      sql.Add('SHOW BINLOG EVENTS IN '+ QuotedStr(pSucursal.Id_archivo_log_procesado) +' LIMIT 500000'); // Ok

      //sql.Add('SHOW BINLOG EVENTS IN '+ QuotedStr(pSucursal.Id_archivo_log_procesado) +' LIMIT 5000');
      //sql.Add('SELECT * FROM bin_log_events_tabla WHERE bin_log_events_tabla.Log_name= '+ QuotedStr(pSucursal.Id_archivo_log_procesado) +' LIMIT 10000');
    end;//end with qrySelect_Binlog_events do
  end
  else
  begin
    with qrySelect_Binlog_events do
    begin
      //sql.Add('SHOW BINLOG EVENTS IN '+  QuotedStr(pSucursal.Id_archivo_log_procesado) + 'FROM '+ IntToStr(pSucursal.Pos_log_procesado) +' LIMIT 400000');
      sql.Add('SHOW BINLOG EVENTS IN '+  QuotedStr(pSucursal.Id_archivo_log_procesado) + 'FROM '+ IntToStr(pSucursal.Pos_log_procesado) +' LIMIT 500000'); // Ok

        //sql.Add('SHOW BINLOG EVENTS IN '+  QuotedStr(pSucursal.Id_archivo_log_procesado) + 'FROM '+ IntToStr(pSucursal.Pos_log_procesado) +' LIMIT 5000'); //
      //sql.Add('SHOW BINLOG EVENTS IN '+  QuotedStr(pSucursal.Id_archivo_log_procesado) + 'FROM '+ IntToStr(pSucursal.Pos_log_procesado) +' LIMIT 50000'); // 50000
      //sql.Add('SELECT * FROM bin_log_events_tabla WHERE bin_log_events_tabla.Log_name='+ QuotedStr(pSucursal.Id_archivo_log_procesado) + ' AND bin_log_events_tabla.Pos >='+ IntToStr(pSucursal.Pos_log_procesado) +' LIMIT 10000');
    end;//end with qrySelect_Binlog_events do

  end;//end if bPrimer_archivo_log = True then
  try
    qrySelect_Binlog_events.Open;
  except
    on E: Exception do
    begin
      sCadena_error:= DateTimeToStr(Now)+'  2---Sucursal:'+IntToStr(pSucursal.Id_sucursal)+' - '+pSucursal.Des_sucursal+' -  Error: '+ (E.Message);

      //pSucursal.Exceptuado:= True;
      if pSucursal.GraboError= False then
        _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
      pSucursal.GraboError:= True;
      _CargarItemsMenu;
      if Assigned (slSentencias_comunicacion) then
        slSentencias_comunicacion.Free;
      Exit;
    end;
  end;

  with qrySelect_Binlog_events do
  begin
    if RecordCount > 0 then
    begin

      bFin_archivo:= False;
      // Es un log Corto de 4 líneas

//      sCadena_error:= 'Id_archivo:'+  pSucursal.Id_archivo_log_procesado + ' Pos: '+ IntToStr(pSucursal.Pos_log_procesado) + ' Recordcount:'+ IntToStr(RecordCount);
//      _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);

      if ((pSucursal.Pos_log_procesado=0) and (RecordCount<=4))  then  // Es un log Corto de 4 líneas
      begin

        if pSucursal.Id_archivo_log_procesado < sUltimo_archivo_log_BDorigen then
        begin
          // Cambio de Archivo
          bFin_archivo:= True;

          pSucursal.Pos_log_procesado:= 0;
          iNro_Extension:= StrToInt(RightStr(pSucursal.Id_archivo_log_procesado, oParametro.Comu_largo_extension_file_binlog)) + 1;
          pSucursal.Id_archivo_log_procesado:= oParametro.Comu_nombre_file_binlog + '.'+ RightStr('00000000000000000'+IntToStr(iNro_Extension), oParametro.Comu_largo_extension_file_binlog);

          //
        end;
      end
      else
      begin

        First;
        while ((Not Eof) and (pSucursal.ErrorVerificacionSentencias>=0)) do  // Se agrega la condicion pSucursal.ErrorVerificacionSentencias<>0 para que reinicie todo de nuevo
        begin

          // Nueva forma de sacar el End_log_pos porque en mariadb windows viene en 0 cuando no es el final de la transaccion
          if FieldByName('End_log_pos').AsInteger>0 then
            iEnd_log_pos_real:= FieldByName('End_log_pos').AsInteger
          else
          begin
            if RecNo+1 <= RecordCount then
            begin
              Next;
              iEnd_log_pos_real:= FieldByName('pos').AsInteger;
              Prior;
            end
            else
              iEnd_log_pos_real:= FieldByName('pos').AsInteger;
          end;
          //
          iEnd_log_pos:=iEnd_log_pos_real;


          bEsAnnotate_rows:= False;
          if ((((PosEx('UPDATE `', FieldByName('info').AsString)>0)      and (PosEx('`articulos` SET', FieldByName('info').AsString)>0))
              or ((PosEx('INSERT INTO `', FieldByName('info').AsString)>0) and (PosEx('`articulos`(', FieldByName('info').AsString)>0))
              or ((PosEx('DELETE FROM `', FieldByName('info').AsString)>0) and (PosEx('`articulos`', FieldByName('info').AsString)>0))) and ( (LowerCase(FieldByName('Event_type').AsString) = 'annotate_rows')) and (bBD_erronea= False)) then // Esto es por la actualizacion Masiva de precios
          begin
            bEsAnnotate_rows:= True;
          end;

          //Event_type en el MySql me identifica de que tipo de sentencia se trata si es = Xid se trata de un commit si es = query se trata de un insert o update o delete
          if (((LeftStr(FieldByName('info').AsString, 6 ) = 'COMMIT') and ( (LowerCase(FieldByName('Event_type').AsString) = 'xid') or (LowerCase(FieldByName('Event_type').AsString) = 'query') or (LowerCase(FieldByName('Event_type').AsString) = 'annotate_rows')) and (bBD_erronea= False)) or bEsAnnotate_rows) then
          begin
            try
              //Agrego al ser unica linea y cerrar la transaccion
              if bEsAnnotate_rows then // Esto es por la actualizacion Masiva de precios
              begin
                sCadena:= FieldByName('info').AsString;
                if ((PosEx('update', LowerCase(sCadena))>0) and (PosEx('set', LowerCase(sCadena))>0)) then
                begin
                  iPosicion_caracter_inicial:=PosEx('update', LowerCase(sCadena));
                  if iPosicion_caracter_inicial> 0  then
                  begin
                    sCadena:= RightStr(sCadena,(Length(sCadena) - (iPosicion_caracter_inicial-1)) );
                    if PosEx('set', LowerCase(sCadena))> 0  then
                    begin
                      // Elimina \'
                      sCadena:= AnsiReplaceStr(sCadena,'\'+chr(39), ' ');
                      _Agregar_CadenaSQL(pSucursal, sCadena, 7, pTabla_Exportar);
                      // Pongo el iPosIni como el pos
                      iPosIni := FieldByName('pos').AsInteger;
                      //
                    end;
                  end;
                end;

                if ((PosEx('insert', LowerCase(sCadena))> 0) and (PosEx('into', LowerCase(sCadena))> 0 ))then
                begin
                  sCadena:= RightStr(sCadena,(Length(sCadena) - (PosEx('insert', LowerCase(sCadena))-1)) );
                  iPosicion_caracter_inicial:=PosEx('into', LowerCase(sCadena));
                  if iPosicion_caracter_inicial> 0  then
                  begin
                    iPosicion_caracter_final:= iPosicion_caracter_inicial + 4;
                    // Elimina \'
                    sCadena:= AnsiReplaceStr(sCadena,'\'+chr(39), ' ');
                    _Agregar_CadenaSQL(pSucursal,sCadena, iPosicion_caracter_final, pTabla_Exportar);
                    // Pongo el iPosIni como el pos
                    iPosIni := FieldByName('pos').AsInteger;
                  end;//end if iPosicion_caracter_inicial> 0  then
                end; //end if  PosEx('insert', sCadena)> 0  then

                // if ((PosEx('delete', LowerCase(sCadena))>0) and (PosEx('where', LowerCase(sCadena))>0)) then
                if ((PosEx('delete', LowerCase(sCadena))>0) and (PosEx('from', LowerCase(sCadena))>0)) then
                begin
                  sCadena:= RightStr(sCadena,(Length(sCadena) - (PosEx('delete', LowerCase(sCadena))-1)) );


                  iPosicion_caracter_inicial:=PosEx('from', LowerCase(sCadena));
                  if iPosicion_caracter_inicial> 0  then
                  begin
                    iPosicion_caracter_final:= iPosicion_caracter_inicial + 4 + 1;
                    // Elimina \'
                    sCadena:= AnsiReplaceStr(sCadena,'\'+chr(39), ' ');
                    _Agregar_CadenaSQL(pSucursal, sCadena, iPosicion_caracter_final, pTabla_Exportar);
                    // Pongo el iPosIni como el pos
                    iPosIni := FieldByName('pos').AsInteger;
                  end;//end if iPosicion_caracter_inicial> 0  then
                end;

              end;
              //

              if ((slSentencias_comunicacion.Count > 0) and (bComprobanteTemporal=False)) then
              begin
                // Nuevo enfoque 1.
                // Verifico si el conjunto de sentencias fue generado por la sucursal a la cual quiero transmitir
                if iPosIni > 0 then
                begin
                  bHayRegistros:= False;
                  qryLog_comunicacion_externo.Close;
                  qryLog_comunicacion_externo.Parameters.ParamByName('p_origen_id_sucursal').Value       := pSucursal.Id_sucursal;
                  qryLog_comunicacion_externo.Parameters.ParamByName('pId_archivo_log_procesado').Value  := FieldByName('Log_name').AsString;
                  qryLog_comunicacion_externo.Parameters.ParamByName('pPos_log_procesado').Value         := iPosIni;

                  //_Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, IntToStr(pSucursal.Id_sucursal)+'  -  '+ FieldByName('Log_name').AsString + '  -  '+ IntToStr(iPosIni));
                  //slSentencias_comunicacion.SaveToFile('c:\vs-comunicacion\sentencias_'+trim(FieldByName('Log_name').AsString+'____'+FieldByName('End_log_pos').AsString)+'.txt');


                  //-------------------------
                  // ojo pv
                  // Version para sucursales
                  // Comentar cuando no lo sea
                  // Para Sucursal: Sin comentar
                  // Para Central: Comentar
                  //-------------------------
//                   qryLog_comunicacion_externo.Parameters.ParamByName('pPos_log_procesado').Value         := 0;
                  //-------------------------

                  qryLog_comunicacion_externo.Open;
                  if qryLog_comunicacion_externo.RecordCount > 0 then   // Si es >0 - No transmitir a esta sucursal porque fue la sucursal origen de las sentencias
                    bHayRegistros:= True
                  else
                    bHayRegistros:= False;

//                  // No transmitir las actualizaciones para el regitro de hora a otras sucursales
//                  for i:= 0 to slSentencias_comunicacion.Count -1 do
//                  begin
//                    if (LeftStr(slSentencias_comunicacion.Strings[i],80)='UPDATE sucursales_log_tiempo_comunicacion SET sucursales_log_tiempo_comunicacion') then
//                    begin
//                      bHayRegistros:= True;
//                      Break;
//                    end;
//                  end;
//                  //

//                  // Verifica si esta la sentencia INSERT INTO log_comunicacion_externo (_origen_id_sucursal, id_archivo_log_procesado, pos_log_procesado) VALUE
//                  // Saltea el lote
//                  // No tengo registros locales en log_comunicacion_externo
//                  for i:= 0 to slSentencias_comunicacion.Count -1 do
//                  begin
//                    if (LeftStr(slSentencias_comunicacion.Strings[i],110)='INSERT INTO log_comunicacion_externo (_origen_id_sucursal, id_archivo_log_procesado, pos_log_procesado) VALUES') then
//                    begin
//                      bHayRegistros:= True;
//                      Break;
//                    end;
//                  end;
//                  //

                  if bHayRegistros = False then
                  begin

                    // slSentencias_comunicacion.SaveToFile('c:\vs-comunicacion\sentencias_'+trim(FieldByName('Log_name').AsString+'____'+FieldByName('End_log_pos').AsString)+'.txt');

                    // Verifico si no pudo verificar la sentencia antes. No debo trasmitir porque ya lo hizo antes
                    if pSucursal.ErrorVerificacionSentencias = 0 then
                    begin
                      qrySentencias_comunicacion.Connection.BeginTrans;
                      //-------------------------
                      // ojo
                      // Version para no controlar claves ni referencias
                      // Comentar cuando no sea necesario
                      //-------------------------
                      //qrySentencias_comunicacion.SQL.Add('SET FOREIGN_KEY_CHECKS=0');
                      //qrySentencias_comunicacion.ExecSQL;
                      //qrySentencias_comunicacion.SQL.Clear;
                      //-------------------------
                    end;

                    // Agrego en el log de la base destino la información necesaria para identificar el orinen (sucursal) de las transacciones
                    iSentencia_Dep_Texto_Enriquecido:=1;
                    SetLength(ArrSentencias_Dep_Texto_Enriquecido, iSentencia_Dep_Texto_Enriquecido);
                    ArrSentencias_Dep_Texto_Enriquecido[iSentencia_Dep_Texto_Enriquecido-1]:=-1;

                    for i:= 0 to slSentencias_comunicacion.Count -1 do
                    begin
                      // Verifico si es actualizacion de Normativa para hacerlo diferenciado
                      if ((LeftStr(slSentencias_comunicacion.Strings[i],45)='UPDATE articulos_gestion SET caracteristicas=')
                        or (LeftStr(slSentencias_comunicacion.Strings[i],66)='INSERT INTO articulos_gestion (id_articulo,caracteristicas) VALUES')
                        or (LeftStr(slSentencias_comunicacion.Strings[i],45)='UPDATE planes_entidad_convenio SET normativa=')) then
                      begin

                        //
                        if LeftStr(slSentencias_comunicacion.Strings[i],45)='UPDATE planes_entidad_convenio SET normativa=' then
                        begin
                          iPosicion_caracter_inicial:=47;
                          iPosicion_caracter_final:= PosEx(chr(39)+' WHERE (id_cliente=', slSentencias_comunicacion.Strings[i]);
                          if iPosicion_caracter_final> 0 then
                          begin
                            //iSentencia_Dep_Texto_Enriquecido:= i;
                            // Cargo en el array de sentencias de texto enriquecido
                            ArrSentencias_Dep_Texto_Enriquecido[iSentencia_Dep_Texto_Enriquecido-1]:=i;
                            // Creo otra posicion del array para la proxima sentencia de texto enriquecido
                            iSentencia_Dep_Texto_Enriquecido:=iSentencia_Dep_Texto_Enriquecido + 1;
                            SetLength(ArrSentencias_Dep_Texto_Enriquecido, iSentencia_Dep_Texto_Enriquecido);
                            // Le asigno -1. Bandera que no se utilizo la posicion todavia
                            ArrSentencias_Dep_Texto_Enriquecido[iSentencia_Dep_Texto_Enriquecido-1]:=-1;
                            //

                            // Verifico si no pudo verificar la sentencia antes. No debo trasmitir porque ya lo hizo antes
                            if pSucursal.ErrorVerificacionSentencias = 0 then
                            begin
                              sTextoEnriquecido:= copy(slSentencias_comunicacion.Strings[i], iPosicion_caracter_inicial, (iPosicion_caracter_final-iPosicion_caracter_inicial));
                              sWhere:= copy(slSentencias_comunicacion.Strings[i],iPosicion_caracter_final+1, (Length(slSentencias_comunicacion.Strings[i])-iPosicion_caracter_final+1));
                              sSqlOriginal:= qryUpdateNormativa.SQL[0];
                              qryUpdateNormativa.SQL[0]:= qryUpdateNormativa.SQL[0]+ sWhere;
                              sTextoEnriquecido:= _Depurar_Texto_Erriquecido(sTextoEnriquecido);
                              qryUpdateNormativa.Parameters.ParamByName('pNormativa').Value := sTextoEnriquecido;
                              //_Grabar_LogErrores_Comunicacion(0, 'Normativa='+sTextoEnriquecido);
                              //_Grabar_LogErrores_Comunicacion(0, 'Where    ='+sWhere);
                              //_Grabar_LogErrores_Comunicacion(0, 'SQL      ='+qryUpdateNormativa.SQL[0]);

                              qryUpdateNormativa.ExecSQL;
                              // Vuelvo a la sentencia original
                              qryUpdateNormativa.SQL[0]:= sSqlOriginal;
                            end;
                          end;//end if iPosicion_caracter_final> 0 then
                        end;

                        //

                        if LeftStr(slSentencias_comunicacion.Strings[i],45)='UPDATE articulos_gestion SET caracteristicas=' then
                        begin
                          iPosicion_caracter_inicial:=47;
                          iPosicion_caracter_final:= PosEx(chr(39)+' WHERE id_articulo=', slSentencias_comunicacion.Strings[i]);
                          if iPosicion_caracter_final> 0 then
                          begin
                            //iSentencia_Dep_Texto_Enriquecido:= i;
                            // Cargo en el array de sentencias de texto enriquecido
                            ArrSentencias_Dep_Texto_Enriquecido[iSentencia_Dep_Texto_Enriquecido-1]:=i;
                            // Creo otra posicion del array para la proxima sentencia de texto enriquecido
                            iSentencia_Dep_Texto_Enriquecido:=iSentencia_Dep_Texto_Enriquecido + 1;
                            SetLength(ArrSentencias_Dep_Texto_Enriquecido, iSentencia_Dep_Texto_Enriquecido);
                            // Le asigno -1. Bandera que no se utilizo la posicion todavia
                            ArrSentencias_Dep_Texto_Enriquecido[iSentencia_Dep_Texto_Enriquecido-1]:=-1;
                            //

                            // Verifico si no pudo verificar la sentencia antes. No debo trasmitir porque ya lo hizo antes
                            if pSucursal.ErrorVerificacionSentencias = 0 then
                            begin
                              sTextoEnriquecido:= copy(slSentencias_comunicacion.Strings[i], iPosicion_caracter_inicial, (iPosicion_caracter_final-iPosicion_caracter_inicial));
                              sWhere:= copy(slSentencias_comunicacion.Strings[i],iPosicion_caracter_final+1, (Length(slSentencias_comunicacion.Strings[i])-iPosicion_caracter_final+1));
                              sSqlOriginal:= qryUpdateArticuloGestion.SQL[0];
                              qryUpdateArticuloGestion.SQL[0]:= qryUpdateArticuloGestion.SQL[0]+ sWhere;
                              sTextoEnriquecido:= _Depurar_Texto_Erriquecido(sTextoEnriquecido);
                              qryUpdateArticuloGestion.Parameters.ParamByName('pCaracteristicas').Value := sTextoEnriquecido;
                              //_Grabar_LogErrores_Comunicacion(0, 'Normativa='+sTextoEnriquecido);
                              //_Grabar_LogErrores_Comunicacion(0, 'Where    ='+sWhere);
                              //_Grabar_LogErrores_Comunicacion(0, 'SQL      ='+qryUpdateNormativa.SQL[0]);

                              qryUpdateArticuloGestion.ExecSQL;
                              // Vuelvo a la sentencia original
                              qryUpdateArticuloGestion.SQL[0]:= sSqlOriginal;
                            end;
                          end;//end if iPosicion_caracter_final> 0 then
                        end;

                        //

                        {
                        if LeftStr(slSentencias_comunicacion.Strings[i],66)='INSERT INTO articulos_gestion (id_articulo,caracteristicas) VALUES' then
                        begin
                          iPosicion_caracter_inicial:=47;
                          iPosicion_caracter_final:= PosEx(chr(39)+' WHERE id_articulo=', slSentencias_comunicacion.Strings[i]);
                          if iPosicion_caracter_final> 0 then
                          begin
                            //iSentencia_Dep_Texto_Enriquecido:= i;
                            // Cargo en el array de sentencias de texto enriquecido
                            ArrSentencias_Dep_Texto_Enriquecido[iSentencia_Dep_Texto_Enriquecido-1]:=i;
                            // Creo otra posicion del array para la proxima sentencia de texto enriquecido
                            iSentencia_Dep_Texto_Enriquecido:=iSentencia_Dep_Texto_Enriquecido + 1;
                            SetLength(ArrSentencias_Dep_Texto_Enriquecido, iSentencia_Dep_Texto_Enriquecido);
                            // Le asigno -1. Bandera que no se utilizo la posicion todavia
                            ArrSentencias_Dep_Texto_Enriquecido[iSentencia_Dep_Texto_Enriquecido-1]:=-1;
                            //

                            // Verifico si no pudo verificar la sentencia antes. No debo trasmitir porque ya lo hizo antes
                            if pSucursal.ErrorVerificacionSentencias = 0 then
                            begin
                              sTextoEnriquecido:= copy(slSentencias_comunicacion.Strings[i], iPosicion_caracter_inicial, (iPosicion_caracter_final-iPosicion_caracter_inicial));
                              sWhere:= copy(slSentencias_comunicacion.Strings[i],iPosicion_caracter_final+1, (Length(slSentencias_comunicacion.Strings[i])-iPosicion_caracter_final+1));
                              sSqlOriginal:= qryUpdateArticuloGestion.SQL[0];
                              qryUpdateArticuloGestion.SQL[0]:= qryUpdateArticuloGestion.SQL[0]+ sWhere;
                              sTextoEnriquecido:= _Depurar_Texto_Erriquecido(sTextoEnriquecido);
                              qryUpdateArticuloGestion.Parameters.ParamByName('pCaracteristicas').Value := sTextoEnriquecido;
                              //_Grabar_LogErrores_Comunicacion(0, 'Normativa='+sTextoEnriquecido);
                              //_Grabar_LogErrores_Comunicacion(0, 'Where    ='+sWhere);
                              //_Grabar_LogErrores_Comunicacion(0, 'SQL      ='+qryUpdateNormativa.SQL[0]);

                              qryUpdateArticuloGestion.ExecSQL;
                              // Vuelvo a la sentencia original
                              qryUpdateArticuloGestion.SQL[0]:= sSqlOriginal;
                            end;
                          end;//end if iPosicion_caracter_final> 0 then
                        end;
                        }

                      end
                      else
                      begin
                        // Verifico si no pudo verificar la sentencia antes. No debo trasmitir porque ya lo hizo antes
                        if pSucursal.ErrorVerificacionSentencias = 0 then
                        begin
                          qrySentencias_comunicacion.SQL.Add(slSentencias_comunicacion.Strings[i]);
                          qrySentencias_comunicacion.ExecSQL;

                          //_Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, slSentencias_comunicacion.Strings[i] + ' 1 ');

                          qrySentencias_comunicacion.SQL.Clear;


                        end;
                      end;
                    end;

                    qrySelect_LastBinaryLogDestino.Close;
                    qrySelect_LastBinaryLogDestino.Open;
                    if qrySelect_LastBinaryLogDestino.RecordCount > 0 then
                    begin



                      if pSucursal.ErrorVerificacionSentencias = 0 then
                      begin
                        pSucursal.Ultimo_archivo_log_BDestino:= qrySelect_LastBinaryLogDestino.FieldByName('file').AsString;
                        pSucursal.Ultimo_Pos_log_BDestino    := qrySelect_LastBinaryLogDestino.FieldByName('position').AsInteger;
                      end;
                      qrySelect_LastBinaryLogDestino.Close;

                      // Verifico si no pudo verificar la sentencia antes. No debo trasmitir porque ya lo hizo antes
                      if pSucursal.ErrorVerificacionSentencias = 0 then
                      begin
                        //-------------------------
                        // ojo
                        // Version para no controlar claves ni referencias
                        // Comentar cuando no sea necesario
                        //-------------------------
                        //qrySentencias_comunicacion.SQL.Add('SET FOREIGN_KEY_CHECKS=1');
                        //qrySentencias_comunicacion.ExecSQL;
                        //qrySentencias_comunicacion.SQL.Clear;
                        //-------------------------

                        sCadena:= 'INSERT INTO log_comunicacion_externo (_origen_id_sucursal, id_archivo_log_procesado, pos_log_procesado) VALUES ('+IntToStr(oParametro.Comu_Id_Sucursal)+', '+QuotedStr(pSucursal.Ultimo_archivo_log_BDestino)+', '+IntToStr(pSucursal.Ultimo_Pos_log_BDestino)+')';
                        qrySentencias_comunicacion.SQL.Add(sCadena);
                        qrySentencias_comunicacion.ExecSQL;

                        //_Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena + ' 2 ');

                        qrySentencias_comunicacion.SQL.Clear;

                        // Actualizo sucursales con el ultimo pos del origen en la tabla sucursal de la BD que recibe la comunicacion

                        // Nueva forma de sacar el End_log_pos porque en mariadb windows viene en 0 cuando no es el final de la transaccion
                        if FieldByName('End_log_pos').AsInteger>0 then
                          iEnd_log_pos_real:= FieldByName('End_log_pos').AsInteger
                        else
                        begin
                          if RecNo+1 <= RecordCount then
                          begin
                            Next;
                            iEnd_log_pos_real:= FieldByName('pos').AsInteger;
                            Prior;
                          end
                          else
                            iEnd_log_pos_real:= FieldByName('pos').AsInteger;
                        end;
                        //

                        sCadena:= 'UPDATE sucursales_log_recepcion SET id_archivo_log_procesado_origen='+QuotedStr(FieldByName('Log_name').AsString)+', pos_log_procesado_origen='+IntToStr(iEnd_log_pos_real)+ ' WHERE id_sucursal='+IntToStr(oParametro.Comu_Id_Sucursal);
                        qrySentencias_comunicacion.SQL.Add(sCadena);
                        qrySentencias_comunicacion.ExecSQL;

                        //_Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena + ' 3 ');

                        qrySentencias_comunicacion.SQL.Clear;

                        qrySentencias_comunicacion.Connection.CommitTrans;
                      end;

                      iUltimo_Pos_log_BDestinoReal:=_Get_Pos_log_procesado_Grabado_Destino(slSentencias_comunicacion, pSucursal.Ultimo_archivo_log_BDestino, pSucursal.Ultimo_Pos_log_BDestino, ArrSentencias_Dep_Texto_Enriquecido);

                      if iUltimo_Pos_log_BDestinoReal=-1 then
                      begin
                        // Si hay Update o delete solamente, osea que no hay Insert > puede ser que no tenga impacto en la BD destino y por ende no guarde en el log de la BD destino.
                        // Entonces no marco como error
                        bHayInsert:= False;
                        for i := 0 to slSentencias_comunicacion.Count-1 do
                        begin
                          if ((PosEx('insert', LowerCase(slSentencias_comunicacion[i]))> 0) and (PosEx('into', LowerCase(slSentencias_comunicacion[i]))> 0 ))then
                          begin
                            bHayInsert:= True;
                            Break;
                          end;
                        end;
                        if bHayInsert=False then
                          iUltimo_Pos_log_BDestinoReal := pSucursal.Ultimo_Pos_log_BDestino;
                        //
                      end;

                      if iUltimo_Pos_log_BDestinoReal >=0 then
                      begin
                        pSucursal.ErrorVerificacionSentencias:= 0;
                        if iUltimo_Pos_log_BDestinoReal<>pSucursal.Ultimo_Pos_log_BDestino then
                        begin
                          // Actualizo con el que grabo realmente
                          sCadena:= 'UPDATE log_comunicacion_externo SET pos_log_procesado='+IntToStr(iUltimo_Pos_log_BDestinoReal)+' WHERE _origen_id_sucursal='+IntToStr(oParametro.Comu_Id_Sucursal)+' AND id_archivo_log_procesado='+QuotedStr(pSucursal.Ultimo_archivo_log_BDestino)+' AND pos_log_procesado='+IntToStr(pSucursal.Ultimo_Pos_log_BDestino);
                          qrySentencias_comunicacion.SQL.Add(sCadena);
                          qrySentencias_comunicacion.Connection.BeginTrans;
                          qrySentencias_comunicacion.ExecSQL;
                          qrySentencias_comunicacion.Connection.CommitTrans;
                          qrySentencias_comunicacion.SQL.Clear;

                          pSucursal.Ultimo_Pos_log_BDestino:= iUltimo_Pos_log_BDestinoReal;
                          //sCadena_error:= DateTimeToStr(Now)+'  6---Sucursal:'+IntToStr(pSucursal.Id_sucursal)+' - '+pSucursal.Des_sucursal+' -  Actualizo Pos destino en log_comunicacion_externo: '+sCadena;
                          //_Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);

                        end;
                      end
                      else
                      begin
                        // En caso de error -1 o -2 no para la comunicacion para la sucursal.
                        if iUltimo_Pos_log_BDestinoReal=-1 then
                          sCadena_error:= DateTimeToStr(Now)+'  7---Sucursal:(-1)'+IntToStr(pSucursal.Id_sucursal)+' - '+pSucursal.Des_sucursal+' -  No encontro sentencias en destino. Pos: '+pSucursal.Ultimo_archivo_log_BDestino+' '+IntToStr(pSucursal.Ultimo_Pos_log_BDestino)+chr(13)+' - Pos origen: '+FieldByName('Log_name').AsString+' '+IntToStr(iPosIni)
                        else if iUltimo_Pos_log_BDestinoReal=-2 then
                          sCadena_error:= DateTimeToStr(Now)+'  8---Sucursal:(-2)'+IntToStr(pSucursal.Id_sucursal)+' - '+pSucursal.Des_sucursal+' -  No pudo ejecutar sentencias SHOW BINLOG EVENT en destino. Pos: '+pSucursal.Ultimo_archivo_log_BDestino+' '+IntToStr(pSucursal.Ultimo_Pos_log_BDestino)+chr(13)+' - Pos origen: '+FieldByName('Log_name').AsString+' '+IntToStr(iPosIni)
                        else if iUltimo_Pos_log_BDestinoReal=-3 then
                          sCadena_error:= DateTimeToStr(Now)+'  7---Sucursal:(-3)'+IntToStr(pSucursal.Id_sucursal)+' - '+pSucursal.Des_sucursal+' -  No encontro sentencias en destino. Pos: '+pSucursal.Ultimo_archivo_log_BDestino+' '+IntToStr(pSucursal.Ultimo_Pos_log_BDestino)+chr(13)+' - Pos origen: '+FieldByName('Log_name').AsString+' '+IntToStr(iPosIni)
                        else if iUltimo_Pos_log_BDestinoReal=-10 then // no entro nada en el registro de eventos de la BD destino
                          sCadena_error:= DateTimeToStr(Now)+'  7---Sucursal:(-10)'+IntToStr(pSucursal.Id_sucursal)+' - '+pSucursal.Des_sucursal+' -  No encontro sentencias en destino. Pos: '+pSucursal.Ultimo_archivo_log_BDestino+' '+IntToStr(pSucursal.Ultimo_Pos_log_BDestino)+chr(13)+' - Pos origen: '+FieldByName('Log_name').AsString+' '+IntToStr(iPosIni);

                        if iUltimo_Pos_log_BDestinoReal<>-10 then
                        begin
                          pSucursal.ErrorVerificacionSentencias:= iUltimo_Pos_log_BDestinoReal;

                          //if pSucursal.GraboError= False then
                            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
                          pSucursal.GraboError:= True;
                          _CargarItemsMenu;
                          if Assigned (slSentencias_comunicacion) then
                            slSentencias_comunicacion.Clear;
                          qrySentencias_comunicacion.SQL.Clear;
                        end
                        else
                        begin
                          //Fuerzo a seguir...
                          // -10 no hubo sentencias. Osea vacio en el log de eventos destino
                        end;
                        //Exit;
                      end;

                      // Si pudo verificar las sentencias en la BD destino: grabo en el log local
                      if pSucursal.ErrorVerificacionSentencias = 0 then
                      begin
                        try

                          // Nueva forma de sacar el End_log_pos porque en mariadb windows viene en 0 cuando no es el final de la transaccion
                          if FieldByName('End_log_pos').AsInteger>0 then
                            iEnd_log_pos_real:= FieldByName('End_log_pos').AsInteger
                          else
                          begin
                            if RecNo+1 <= RecordCount then
                            begin
                              Next;
                              iEnd_log_pos_real:= FieldByName('pos').AsInteger;
                              Prior;
                            end
                            else
                              iEnd_log_pos_real:= FieldByName('pos').AsInteger;
                          end;
                          //

                          qryUpdate_FileandPos_BDDestino.Parameters.ParamByName('pId_archivo_log_procesado').Value  := FieldByName('Log_name').AsString;
                          qryUpdate_FileandPos_BDDestino.Parameters.ParamByName('pPos_log_procesado').Value         := iEnd_log_pos_real;
                          qryUpdate_FileandPos_BDDestino.Parameters.ParamByName('pId_sucursal').Value               := pSucursal.Id_sucursal;

                          qryUpdate_FileandPos_BDDestino.Connection.BeginTrans;
                          qryUpdate_FileandPos_BDDestino.ExecSQL;

                          qryUpdate_FileandPos_BDDestino.Connection.CommitTrans;

                          //Actualizo el objeto sucursal que tengo en memoria con los ultimos Log_name y pos que logre actualizar la BD destino
                          //Blanqueo la variable de grabacion del error
                          pSucursal.Id_archivo_log_procesado  := FieldByName('Log_name').AsString;
                          pSucursal.Pos_log_procesado         := iEnd_log_pos_real;
                          pSucursal.GraboError                := False;
                          pSucursal.GraboError2Instancia      := False;

                          //_Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, 'Verifico' + ' 4 ');

                        except
                          on E: Exception do
                          begin
                            sCadena_error:= DateTimeToStr(Now)+'  10---Sucursal:'+IntToStr(pSucursal.Id_sucursal)+' - '+pSucursal.Des_sucursal+' -  Error: '+ (E.Message)+
                                                                Chr(13)+'SQL: '+ qryUpdate_FileandPos_BDDestino.SQL.Text;

                            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);

                            //qryUpdate_FileandPos_BDDestino.Connection.RollbackTrans;

                            //Exit;
                          end;
                        end;//end try}
                      end;

                    end
                    else
                    begin
                      // Desbloqueo la BD
                      //sCadena:= 'UNLOCK TABLES';
                      //qrySentencias_comunicacion.SQL.Add(sCadena);
                      //qrySentencias_comunicacion.ExecSQL;
                      //qrySentencias_comunicacion.SQL.Clear;
                      //

                      // Verifico si no pudo verificar la sentencia antes. No debo trasmitir porque ya lo hizo antes
                      if pSucursal.ErrorVerificacionSentencias = 0 then
                      begin
                        // no pudo abrir => Desago la operacion
                        if qrySentencias_comunicacion.Connection.InTransaction then
                          qrySentencias_comunicacion.Connection.RollbackTrans;
                      end;
                    end;
                    qrySentencias_comunicacion.SQL.Clear;
                    // Fin bloque Nuevo enfoque 2
                    //

                  end
                  else
                  begin

                    try

                      // Nueva forma de sacar el End_log_pos porque en mariadb windows viene en 0 cuando no es el final de la transaccion
                      if FieldByName('End_log_pos').AsInteger>0 then
                        iEnd_log_pos_real:= FieldByName('End_log_pos').AsInteger
                      else
                      begin
                        if RecNo+1 <= RecordCount then
                        begin
                          Next;
                          iEnd_log_pos_real:= FieldByName('pos').AsInteger;
                          Prior;
                        end
                        else
                          iEnd_log_pos_real:= FieldByName('pos').AsInteger;
                      end;
                      //

                      qryUpdate_FileandPos_BDDestino.Parameters.ParamByName('pId_archivo_log_procesado').Value  := FieldByName('Log_name').AsString;
                      qryUpdate_FileandPos_BDDestino.Parameters.ParamByName('pPos_log_procesado').Value         := iEnd_log_pos_real;
                      qryUpdate_FileandPos_BDDestino.Parameters.ParamByName('pId_sucursal').Value               := pSucursal.Id_sucursal;

                      qryUpdate_FileandPos_BDDestino.Connection.BeginTrans;
                      qryUpdate_FileandPos_BDDestino.ExecSQL;

                      qryUpdate_FileandPos_BDDestino.Connection.CommitTrans;

                      //Actualizo el objeto sucursal que tengo en memoria con los ultimos Log_name y pos que logre actualizar la BD destino
                      //Blanqueo la variable de grabacion del error
                      pSucursal.Id_archivo_log_procesado  := FieldByName('Log_name').AsString;
                      pSucursal.Pos_log_procesado         := iEnd_log_pos_real;
                      pSucursal.GraboError                := False;
                      pSucursal.GraboError2Instancia      := False;
                    except
                      on E: Exception do
                      begin
                        sCadena_error:= DateTimeToStr(Now)+'  11---Sucursal:'+IntToStr(pSucursal.Id_sucursal)+' - '+pSucursal.Des_sucursal+' -  Error: '+ (E.Message)+
                                                            Chr(13)+'SQL: '+ qryUpdate_FileandPos_BDDestino.SQL.Text;
                        _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);

                        //qryUpdate_FileandPos_BDDestino.Connection.RollbackTrans;

                        //Exit;
                      end;

                    end;

                  end;
                end
                else
                begin

                end; // if iPosIni > 0 then  // Fin bloque Nuevo enfoque 1
                //
              end
              else
              begin
                try

                  // Nueva forma de sacar el End_log_pos porque en mariadb windows viene en 0 cuando no es el final de la transaccion
                  if FieldByName('End_log_pos').AsInteger>0 then
                    iEnd_log_pos_real:= FieldByName('End_log_pos').AsInteger
                  else
                  begin
                    if RecNo+1 <= RecordCount then
                    begin
                      Next;
                      iEnd_log_pos_real:= FieldByName('pos').AsInteger;
                      Prior;
                    end
                    else
                      iEnd_log_pos_real:= FieldByName('pos').AsInteger;
                  end;
                  //

                  qryUpdate_FileandPos_BDDestino.Parameters.ParamByName('pId_archivo_log_procesado').Value  := FieldByName('Log_name').AsString;
                  qryUpdate_FileandPos_BDDestino.Parameters.ParamByName('pPos_log_procesado').Value         := iEnd_log_pos_real;
                  qryUpdate_FileandPos_BDDestino.Parameters.ParamByName('pId_sucursal').Value               := pSucursal.Id_sucursal;

                  qryUpdate_FileandPos_BDDestino.Connection.BeginTrans;
                  qryUpdate_FileandPos_BDDestino.ExecSQL;

                  qryUpdate_FileandPos_BDDestino.Connection.CommitTrans;

                  //Actualizo el objeto sucursal que tengo en memoria con los ultimos Log_name y pos que logre actualizar la BD destino
                  //Blanqueo la variable de grabacion del error
                  pSucursal.Id_archivo_log_procesado  := FieldByName('Log_name').AsString;
                  pSucursal.Pos_log_procesado         := iEnd_log_pos_real;
                  pSucursal.GraboError                := False;
                  pSucursal.GraboError2Instancia      := False;
                except
                  on E: Exception do
                  begin
                    sCadena_error:= DateTimeToStr(Now)+'  12---Sucursal:'+IntToStr(pSucursal.Id_sucursal)+' - '+pSucursal.Des_sucursal+' -  Error: '+ (E.Message)+
                                                        Chr(13)+'SQL: '+ qryUpdate_FileandPos_BDDestino.SQL.Text;
                    _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);

                    //qryUpdate_FileandPos_BDDestino.Connection.RollbackTrans;

                    //Exit;
                  end;

                end;

              end; //if slSentencias_comunicacion.Count > 0 then
              iPosIni:= -1;
              slSentencias_comunicacion.Clear;
            except
              on E: Exception do
              begin
                if qrySentencias_comunicacion.Connection.InTransaction then
                  qrySentencias_comunicacion.Connection.RollbackTrans;

                sCadena_error:='';
                for i:=0 to slSentencias_comunicacion.Count-1 do
                begin
                  sCadena_error:= sCadena_error + slSentencias_comunicacion.Strings[i] +Chr(13);
                end;

                sCadena_error:= DateTimeToStr(Now)+'  3---Sucursal:'+IntToStr(pSucursal.Id_sucursal)+' - '+pSucursal.Des_sucursal+ ' - Log Name: '+ FieldByName('Log_name').AsString+ ', Pos:'+
                                 IntToStr(iPosIni)+'('+IntToStr(E.HelpContext)+')'+(E.Message)+ Chr(13)+'SQL: '+ sCadena_error;
                                 //Chr(13)+'pSucursal.Tablas_exceptuadas_sucursal: '+ Chr(13)+pSucursal.Tablas_exceptuadas_sucursal.Text;
                                 //FieldByName('End_log_pos').AsInteger

                //[MySQL][ODBC 3.51 Driver][mysqld-5.1.45-log]Duplicate entry 'TB0040-00253133' for key 'PRIMARY'

                // El error salto al querer actualizar en qryUpdate_FileandPos_BDDestino (Tabla sucursal BD Local)
                if qryUpdate_FileandPos_BDDestino.Connection.InTransaction = true then
                begin
                  qryUpdate_FileandPos_BDDestino.Connection.RollbackTrans;
                  sCadena_error:= sCadena_error+chr(13)+DateTimeToStr(Now)+'  3.1---Sucursal:'+IntToStr(pSucursal.Id_sucursal)+' - '+pSucursal.Des_sucursal+' -  Error: No se pudo actualizar en la tabla sucursal de la BD local.'+
                                                              Chr(13)+'SQL: '+ qryUpdate_FileandPos_BDDestino.SQL.Text;
                end;

                //pSucursal.Exceptuado:= true;
                if pSucursal.GraboError= False then
                  _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
                pSucursal.GraboError:= true;
                _CargarItemsMenu;

                //////////////////////////////////////////////////////////
                // Para saltar los duplicados de:
                // comprobantes de ventas
                // Pedidos....
                //////////////////////////////////////////////////////////////////
                if ((PosEx('3---Sucursal:', sCadena_error)> 0) and (PosEx('Duplicate entry', sCadena_error)> 0)
                  and ((PosEx('INSERT INTO comprobantes_de_ventas', sCadena_error)> 0)
                    or (PosEx('INSERT INTO `comprobantes_de_ventas`', sCadena_error)> 0)
                    or (PosEx('INSERT INTO pedidos', sCadena_error)> 0)
                    or (PosEx('INSERT INTO `pedidos`', sCadena_error)> 0)
                    or (PosEx('INSERT INTO comprobantes_de_envios', sCadena_error)> 0)
                    or (PosEx('INSERT INTO `comprobantes_de_envios`', sCadena_error)> 0)
                    or (PosEx('INSERT INTO `registro_de_cajas`', sCadena_error)> 0)
                    or (PosEx('INSERT INTO registro_de_cajas', sCadena_error)> 0)
                    or (PosEx('INSERT INTO `comprobantes_mov_stock`', sCadena_error)> 0)
                    or (PosEx('INSERT INTO comprobantes_mov_stock', sCadena_error)> 0)
                    or (PosEx('INSERT INTO `presentaciones_entidad_convenio`', sCadena_error)> 0)
                    or (PosEx('INSERT INTO presentaciones_entidad_convenio', sCadena_error)> 0)
                    or (PosEx('INSERT INTO `presentaciones_entidad_debito`', sCadena_error)> 0)
                    or (PosEx('INSERT INTO presentaciones_entidad_debito', sCadena_error)> 0)
                    or (PosEx('INSERT INTO `recibos`', sCadena_error)> 0)
                    or (PosEx('INSERT INTO recibos', sCadena_error)> 0)
                    or (PosEx('INSERT INTO `movimientos_caja`', sCadena_error)> 0)
                    or (PosEx('INSERT INTO movimientos_caja', sCadena_error)> 0)
                    or (PosEx('`movimientos_stock` (', sCadena_error)> 0)
                    or (PosEx('movimientos_stock (', sCadena_error)> 0)
                    or (PosEx('INSERT INTO `presentaciones_tarjeta_empresa`', sCadena_error)> 0)
                    or (PosEx('insert into'+ #13 + #10 +'  presentaciones_tarjeta_empresa', sCadena_error)> 0)
                    or (PosEx('INSERT INTO presentaciones_tarjeta_empresa', sCadena_error)> 0)
                    or (PosEx('INSERT INTO'+ #13 + #10 +'personas (', sCadena_error)> 0)
                    or (PosEx('INSERT INTO personas', sCadena_error)> 0)
                    or (PosEx('INSERT INTO'+ #13 + #10 +'`personas`', sCadena_error)> 0)
                    or (PosEx('INSERT INTO `personas`', sCadena_error)> 0)
                    or (PosEx('INSERT INTO'+ #13 + #10 +'comprobantes_de_envio_dom (', sCadena_error)> 0)
                    or (PosEx('INSERT INTO `comprobantes_de_envio_dom`', sCadena_error)> 0)
                    or (PosEx('INSERT INTO comprobantes_de_envio_dom', sCadena_error)> 0)
                    or (PosEx('INSERT INTO'+ #13 + #10 +'comprobantes_de_envio_dom(', sCadena_error)> 0)
                    or ((PosEx('INSERT INTO', LeftStr(sCadena_error,200))> 0) and (PosEx('comprobantes_de_envio_dom', LeftStr(sCadena_error,200))> 0))
                    or (PosEx('INSERT INTO `vendedores`', sCadena_error)> 0)
                    or (PosEx('INSERT INTO vendedores', sCadena_error)> 0)
                    or (PosEx('INSERT INTO `entidad_conv_identificacion`', sCadena_error)> 0)
                    or (PosEx('INSERT INTO entidad_conv_identificacion', sCadena_error)> 0)
                    or (PosEx('INSERT INTO comprobantes_compras', sCadena_error)> 0)
                    or (PosEx('INSERT INTO `comprobantes_compras`', sCadena_error)> 0)
                    or (PosEx('INSERT INTO orden_pago_a_proveedores', sCadena_error)> 0)
                    or (PosEx('INSERT INTO `orden_pago_a_proveedores`', sCadena_error)> 0)
                    or (PosEx('INSERT INTO solicitud_nc_proveedor', sCadena_error)> 0)
                    or (PosEx('INSERT INTO `solicitud_nc_proveedor`', sCadena_error)> 0)
                    or (PosEx('INSERT INTO `cupones_entidad_debito` (`id_cupon`, `id_empresa`, `id_comprobante`,', sCadena_error)> 0)
                    or (PosEx('UPDATE cajas_por_sucursal SET cajas_por_sucursal.z_bi =', sCadena_error)> 0)
                    or (PosEx('UPDATE `cajas_por_sucursal` SET cajas_por_sucursal.z_bi =', sCadena_error)> 0)
                    or (PosEx('INSERT INTO valores_caja (id_empresa, id_sucursal, nro_caja, periodo_caja, id_moneda, id_valor, cantidad)', sCadena_error)> 0)
                    or (PosEx('INSERT INTO `asientos_contables`', sCadena_error)> 0)
                    or (PosEx('INSERT INTO asientos_contables', sCadena_error)> 0)  )) then
                    //      or (PosEx('UPDATE cupones_entidad_convenio', sCadena_error)> 0)
                    //      or (PosEx('INSERT INTO cupones_entidad_debito', sCadena_error)> 0)
                    //      or (PosEx('INSERT INTO auditoria', sCadena_error)> 0)
                    //      or ((PosEx('insert into', sCadena_error)> 0) and (PosEx('presentaciones_entidad_debito', sCadena_error)> 0))
                    //      or ((PosEx('INSERT INTO', sCadena_error)> 0) and (PosEx('movimientos_fondos', sCadena_error)> 0))
                    //      or ((PosEx('INSERT INTO', sCadena_error)> 0) and (PosEx('movimientos_fondos', sCadena_error)> 0))
                    //      or (PosEx('DELETE FROM codigos_barra_articulo WHERE id_articulo', sCadena_error)> 0)
                    //      or (PosEx('INSERT INTO cajas_por_sucursal_fe_caea', sCadena_error)> 0)
                    //      or ((PosEx('INSERT INTO', sCadena_error)> 0) and (PosEx('organizaciones', sCadena_error)> 0))
                    //      or ((PosEx('INSERT INTO', sCadena_error)> 0) and (PosEx('articulos', sCadena_error)> 0))
                    //(PosEx('insert into', LowerCase(sCadena_error))> 0) and (PosEx('presentaciones_tarjeta_empresa', LowerCase(sCadena_error))> 0))
                begin

                  // Nueva forma de sacar el End_log_pos porque en mariadb windows viene en 0 cuando no es el final de la transaccion
                  if FieldByName('End_log_pos').AsInteger>0 then
                    iEnd_log_pos_real:= FieldByName('End_log_pos').AsInteger
                  else
                  begin
                    if RecNo+1 <= RecordCount then
                    begin
                      Next;
                      iEnd_log_pos_real:= FieldByName('pos').AsInteger;
                      Prior;
                    end
                    else
                      iEnd_log_pos_real:= FieldByName('pos').AsInteger;
                  end;
                  //

                  // Salto cuando es duplidado de comprobante
                  qryUpdate_FileandPos_BDDestino.Connection.BeginTrans;

                  qryUpdate_FileandPos_BDDestino.Parameters.ParamByName('pId_archivo_log_procesado').Value  := FieldByName('Log_name').AsString;
                  qryUpdate_FileandPos_BDDestino.Parameters.ParamByName('pPos_log_procesado').Value         := iEnd_log_pos_real;
                  qryUpdate_FileandPos_BDDestino.Parameters.ParamByName('pId_sucursal').Value               := pSucursal.Id_sucursal;
                  qryUpdate_FileandPos_BDDestino.ExecSQL;

                  qryUpdate_FileandPos_BDDestino.Connection.CommitTrans;

                  //Actualizo el objeto sucursal que tengo en memoria con los ultimos Log_name y pos que logre actualizar la BD destino
                  //Blanqueo la variable de grabacion del error
                  pSucursal.Id_archivo_log_procesado  := FieldByName('Log_name').AsString;
                  pSucursal.Pos_log_procesado         := iEnd_log_pos_real;
                  pSucursal.GraboError                := False;
                  pSucursal.GraboError2Instancia      := False;
                  sCadena_error:= '';
                  if Assigned (slSentencias_comunicacion) then
                    slSentencias_comunicacion.Clear;
                  qrySentencias_comunicacion.SQL.Clear;
                  iPosIni:= -1;
                end
                /////////////////////////////////////////////////
                // Inserta comprobante Faltante
                // Pedidos
                /////////////////////////////////////////////////
                else if ((PosEx('3---Sucursal:', sCadena_error)> 0)
                  and (PosEx('Cannot add or update a child row: a foreign key constraint fails', sCadena_error)> 0)
                  and (
                       (PosEx('`lin_comprobantes_de_ventas`, CONSTRAINT `lin_comprobantes_de_ventas_ibfk_3` FOREIGN KEY (`id_pedido`) REFERENCES `pedidos` (`id_pedido`)', sCadena_error)> 0)
                    or (PosEx('`lin_comprobantes_de_ventas`, CONSTRAINT `lin_comprobantes_de_ventas_ibfk_3` FOREIGN KEY (`id_empresa`, `id_pedido`) REFERENCES `pedidos` (`id_empresa`, `id_pedido`)', sCadena_error)> 0)
                    or (PosEx('`lin_pedido`, CONSTRAINT `lin_pedido_ibfk_1` FOREIGN KEY (`id_empresa`, `id_pedido`) REFERENCES `pedidos` (`id_empresa`, `id_pedido`)', sCadena_error)> 0)
                    or (PosEx('`lin_pedido`, CONSTRAINT `lin_pedido_ibfk_1` FOREIGN KEY (`id_pedido`) REFERENCES `pedidos` (`id_pedido`)', sCadena_error)> 0)
                    )) then
                begin
                  if slSentencias_comunicacion.Count > 0 then
                  begin
                    if LeftStr(slSentencias_comunicacion[0], 47) = 'UPDATE lin_comprobantes_de_ventas SET id_pedido' then
                    begin
                      try
                        sSentencia:= slSentencias_comunicacion[0];
                        sId_comprobante:='';
                        iPosicion_caracter_inicial:=PosEx('PE', sSentencia);
                        if PosEx('PE', sSentencia)> 0 then
                        begin
                          sId_comprobante:= copy(sSentencia,iPosicion_caracter_inicial, 15);
                        end;

                        sId_empresa := '0';
                        iPosicion_caracter_inicial:=PosEx('lin_comprobantes_de_ventas.id_empresa =', sSentencia);
                        if iPosicion_caracter_inicial> -1 then
                        begin
                          iPosicion_caracter_inicial := iPosicion_caracter_inicial + 39;
                          sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, Length(sSentencia)-iPosicion_caracter_inicial);
                          iPosicion_caracter_final:=PosEx('AND',sSentencia);
                          if iPosicion_caracter_final> -1 then
                          begin
                            sId_empresa:= Trim(copy(sSentencia,1, (iPosicion_caracter_final-1)));
                          end;
                        end;

                        qrySentencias_comunicacion.SQL.Clear;

                        slSentenciasManuales:= _Confecciona_Sentencias_Faltantes('pedido', sId_empresa, sId_comprobante, '', '',  pSucursal);
                        if Assigned(slSentenciasManuales) then
                        begin
                          if slSentenciasManuales.Count>0 then
                          begin
                            qrySentencias_comunicacion.Connection.BeginTrans;
                            for i:=0 to slSentenciasManuales.Count-1 do
                            begin
                              qrySentencias_comunicacion.SQL.Clear;
                              qrySentencias_comunicacion.SQL.Add(slSentenciasManuales[i]);
                              qrySentencias_comunicacion.ExecSQL;
                            end;

                            qrySentencias_comunicacion.Connection.CommitTrans;

                            sCadena_error:= chr(13) + 'Pedido Insertado y corregido: Empresa:'+sId_empresa+ ' Comprobante: '+ sId_comprobante;
                            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);

                            sCadena_error:= '';
                            if Assigned (slSentencias_comunicacion) then
                              slSentencias_comunicacion.Clear;

                            FreeAndNil(slSentenciasManuales);

                            pSucursal.GraboError           := False;
                            pSucursal.GraboError2Instancia := False;

                            sCadena_error:= '';
                            if Assigned (slSentencias_comunicacion) then
                              slSentencias_comunicacion.Clear;
                          end;
                        end;
                      except
                        on E: Exception do
                        begin

                          if qrySentencias_comunicacion.Connection.InTransaction then
                            qrySentencias_comunicacion.Connection.RollbackTrans;

                          sCadena_error:= 'No se pudo insertar y corregir el Pedido:' + sId_comprobante + Chr(13) + '('+IntToStr(E.HelpContext)+')'+(E.Message)+ Chr(13)+'SQL: ';
                          if Assigned(slSentenciasManuales) then
                          begin
                            for i:=0 to slSentenciasManuales.Count-1 do
                            begin
                              sCadena_error:= sCadena_error + slSentenciasManuales[i] +Chr(13);
                            end;
                            FreeAndNil(slSentenciasManuales);
                          end;
                          if pSucursal.GraboError2Instancia = False then
                          begin
                            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
                            pSucursal.GraboError2Instancia:= true;
                            _CargarItemsMenu;
                          end;
                        end;
                      end;
                      qrySentencias_comunicacion.SQL.Clear;
                    end
                    else if LeftStr(slSentencias_comunicacion[0], 22) = 'DELETE FROM lin_pedido' then
                    begin
                      try
                        sSentencia:= slSentencias_comunicacion[0];
                        sId_comprobante:='';
                        iPosicion_caracter_inicial:=PosEx('PE', sSentencia);
                        if PosEx('PE', sSentencia)> 0 then
                        begin
                          sId_comprobante:= copy(sSentencia,iPosicion_caracter_inicial, 15);
                        end;

                        sId_empresa := '0';
                        iPosicion_caracter_inicial:=PosEx('AND id_empresa =', sSentencia);
                        if iPosicion_caracter_inicial> -1 then
                        begin
                          iPosicion_caracter_inicial := iPosicion_caracter_inicial + 16;
                          sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, Length(sSentencia)-iPosicion_caracter_inicial);
                          //iPosicion_caracter_final:=PosEx('AND',sSentencia);
                          //if iPosicion_caracter_final> -1 then
                          //begin
                            //sId_empresa:= Trim(copy(sSentencia,1, (iPosicion_caracter_final-1)));
                            sId_empresa:= Trim(sSentencia);
                          //end;
                        end;

                        qrySentencias_comunicacion.SQL.Clear;

                        slSentenciasManuales:= _Confecciona_Sentencias_Faltantes('pedido', sId_empresa, sId_comprobante, '', '', pSucursal);
                        if Assigned(slSentenciasManuales) then
                        begin
                          if slSentenciasManuales.Count>0 then
                          begin
                            qrySentencias_comunicacion.Connection.BeginTrans;
                            for i:=0 to slSentenciasManuales.Count-1 do
                            begin
                              qrySentencias_comunicacion.SQL.Clear;
                              qrySentencias_comunicacion.SQL.Add(slSentenciasManuales[i]);
                              qrySentencias_comunicacion.ExecSQL;
                            end;

                            qrySentencias_comunicacion.Connection.CommitTrans;

                            sCadena_error:= chr(13) + 'Pedido Insertado y corregido: Empresa:'+sId_empresa+ ' Comprobante: '+ sId_comprobante;
                            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);

                            sCadena_error:= '';
                            if Assigned (slSentencias_comunicacion) then
                              slSentencias_comunicacion.Clear;

                            FreeAndNil(slSentenciasManuales);

                            pSucursal.GraboError           := False;
                            pSucursal.GraboError2Instancia := False;

                            sCadena_error:= '';
                            if Assigned (slSentencias_comunicacion) then
                              slSentencias_comunicacion.Clear;
                          end;
                        end;
                      except
                        on E: Exception do
                        begin

                          if qrySentencias_comunicacion.Connection.InTransaction then
                            qrySentencias_comunicacion.Connection.RollbackTrans;

                          sCadena_error:= 'No se pudo insertar y corregir el Pedido:' + sId_comprobante + Chr(13) + '('+IntToStr(E.HelpContext)+')'+(E.Message)+ Chr(13)+'SQL: ';
                          if Assigned(slSentenciasManuales) then
                          begin
                            for i:=0 to slSentenciasManuales.Count-1 do
                            begin
                              sCadena_error:= sCadena_error + slSentenciasManuales[i] +Chr(13);
                            end;
                            FreeAndNil(slSentenciasManuales);
                          end;
                          if pSucursal.GraboError2Instancia = False then
                          begin
                            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
                            pSucursal.GraboError2Instancia:= true;
                            _CargarItemsMenu;
                          end;
                        end;
                      end;
                      qrySentencias_comunicacion.SQL.Clear;
                    end;
                  end;
                end

                /////////////////////////////////////////////////
                // Inserta comprobante Faltante
                // Envio
                /////////////////////////////////////////////////
                else if ((PosEx('3---Sucursal:', sCadena_error)> 0)
                  and (PosEx('Cannot add or update a child row: a foreign key constraint fails', sCadena_error)> 0)
                  and (PosEx('`ingreso_mercaderia`, CONSTRAINT `ingreso_mercaderia_ibfk_2` FOREIGN KEY (`id_empresa`, `_envio_id_comprobante`) REFERENCES `comprobantes_de_envios` (`id_empresa`, `id_comprobante`)', sCadena_error)> 0)) then
                begin
                  if slSentencias_comunicacion.Count > 0 then
                  begin

                    // busco la sentencia que contiene el ingreso
                    sId_comprobante:='';
                    for i := 0 to slSentencias_comunicacion.Count-1 do
                    begin
                      if LeftStr(slSentencias_comunicacion[i],30)= 'INSERT INTO ingreso_mercaderia' then
                      begin
                        sSentencia:= slSentencias_comunicacion[i];

                        iPosicion_caracter_inicial:=PosEx('  _latin1'+Chr(39)+'NE', sSentencia);
                        if iPosicion_caracter_inicial>-1 then
                        begin
                          iPosicion_caracter_inicial:= iPosicion_caracter_inicial + 10;
                        end;
                        sId_comprobante:= copy(sSentencia,iPosicion_caracter_inicial, 15);
                        iPosicion_caracter_inicial:=PosEx('VALUES (', sSentencia);
                        if iPosicion_caracter_inicial> -1 then
                        begin
                          iPosicion_caracter_inicial := iPosicion_caracter_inicial + 8;
                          sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, length(sSentencia)-iPosicion_caracter_inicial);
                          iPosicion_caracter_final:=PosEx(',',sSentencia);
                          if iPosicion_caracter_final> -1 then
                          begin
                            sId_empresa:= Trim(copy(sSentencia,1, (iPosicion_caracter_final-1)));
                          end;
                        end;
                        Break;
                      end;
                    end;

                    if ((Length(sId_comprobante)>0) and (Length(sId_empresa)>0)) then
                    begin
                      try

                        qrySentencias_comunicacion.SQL.Clear;

                        slSentenciasManuales:= _Confecciona_Sentencias_Faltantes('envio', sId_empresa, sId_comprobante, '', '', pSucursal);
                        if Assigned(slSentenciasManuales) then
                        begin
                          if slSentenciasManuales.Count>0 then
                          begin
                            qrySentencias_comunicacion.Connection.BeginTrans;
                            for i:=0 to slSentenciasManuales.Count-1 do
                            begin
                              qrySentencias_comunicacion.SQL.Clear;
                              qrySentencias_comunicacion.SQL.Add(slSentenciasManuales[i]);
                              qrySentencias_comunicacion.ExecSQL;
                            end;

                            qrySentencias_comunicacion.Connection.CommitTrans;

                            sCadena_error:= chr(13) + 'Nota de Envio Insertada y corregida: Empresa:'+sId_empresa+ ' Comprobante: '+ sId_empresa+'-'+sId_comprobante;
                            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);

                            sCadena_error:= '';
                            if Assigned (slSentencias_comunicacion) then
                              slSentencias_comunicacion.Clear;

                            FreeAndNil(slSentenciasManuales);

                            pSucursal.GraboError           := False;
                            pSucursal.GraboError2Instancia := False;

                            sCadena_error:= '';
                            if Assigned (slSentencias_comunicacion) then
                              slSentencias_comunicacion.Clear;
                          end;
                        end;
                      except
                        on E: Exception do
                        begin

                          if qrySentencias_comunicacion.Connection.InTransaction then
                            qrySentencias_comunicacion.Connection.RollbackTrans;

                          sCadena_error:= 'No se pudo insertar y corregir la Nota de Envio: Empresa:'+sId_empresa+ ' Comprobante: '+ sId_comprobante + Chr(13) + '('+IntToStr(E.HelpContext)+')'+(E.Message)+ Chr(13)+'SQL: ';
                          if Assigned(slSentenciasManuales) then
                          begin
                            for i:=0 to slSentenciasManuales.Count-1 do
                            begin
                              sCadena_error:= sCadena_error + slSentenciasManuales[i] +Chr(13);
                            end;
                            FreeAndNil(slSentenciasManuales);
                          end;
                          if pSucursal.GraboError2Instancia = False then
                          begin
                            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
                            pSucursal.GraboError2Instancia:= true;
                            _CargarItemsMenu;
                          end;
                        end;
                      end;
                      qrySentencias_comunicacion.SQL.Clear;
                    end
                    else
                    begin
                      if pSucursal.GraboError2Instancia = False then
                      begin
                        sCadena_error:= 'No se pudo insertar y corregir la Nota de Envio: Empresa:'+sId_empresa+ chr(13)+ ' Comprobante: '+ sId_comprobante;
                        _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
                        pSucursal.GraboError2Instancia:= true;
                        _CargarItemsMenu;
                      end;
                    end;
                  end;
                end


                /////////////////////////////////////////////////
                // Inserta comprobante Faltante
                // Factura - Referenciada por NC
                /////////////////////////////////////////////////
                else if ((PosEx('3---Sucursal:', sCadena_error)> 0)
                  and (PosEx('Cannot add or update a child row: a foreign key constraint fails', sCadena_error)> 0)
                  and (PosEx('`notas_credito`, CONSTRAINT `notas_credito_ibfk_2` FOREIGN KEY (`id_empresa`, `_fac_id_comprobante`) REFERENCES `comprobantes_de_ventas` (`id_empresa`, `id_comprobante`)', sCadena_error)> 0)) then
                begin
                  if slSentencias_comunicacion.Count > 0 then
                  begin

                    // busco la sentencia que contiene el ingreso

                    sId_empresa:='';
                    sId_comprobante:='';
                    for i := 0 to slSentencias_comunicacion.Count-1 do
                    begin
                      if LeftStr(slSentencias_comunicacion[i],24)= 'UPDATE cupones_promocion' then
                      begin
                        sSentencia:= slSentencias_comunicacion[i];
                        iPosicion_caracter_inicial:=PosEx('WHERE id_empresa =', sSentencia);
                        if iPosicion_caracter_inicial> -1 then
                        begin
                          iPosicion_caracter_inicial := iPosicion_caracter_inicial + 19;
                          sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, length(sSentencia)-iPosicion_caracter_inicial);
                          iPosicion_caracter_final:=PosEx('AND',sSentencia);
                          if iPosicion_caracter_final> -1 then
                          begin
                            sId_empresa:= Trim(copy(sSentencia,0, (iPosicion_caracter_final-1)));
                          end;
                        end;
                        iPosicion_caracter_inicial:=PosEx('id_comprobante = ', sSentencia);
                        if iPosicion_caracter_inicial>-1 then
                        begin
                          iPosicion_caracter_inicial:= iPosicion_caracter_inicial + 25;
                          sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, length(sSentencia)-iPosicion_caracter_inicial);
                          iPosicion_caracter_final:=PosEx(Chr(39),sSentencia);
                          if iPosicion_caracter_final> -1 then
                          begin
                            sId_comprobante:= Trim(copy(sSentencia,1, (iPosicion_caracter_final-1)));
                          end;
                        end;
                        Break;
                      end;
                    end;

                    if ((Length(sId_comprobante)>0) and (Length(sId_empresa)>0)) then
                    begin
                      try

                        qrySentencias_comunicacion.SQL.Clear;

                        slSentenciasManuales:= _Confecciona_Sentencias_Faltantes('factura', sId_empresa, sId_comprobante, '', '', pSucursal);
                        if Assigned(slSentenciasManuales) then
                        begin
                          if slSentenciasManuales.Count>0 then
                          begin
                            qrySentencias_comunicacion.Connection.BeginTrans;
                            for i:=0 to slSentenciasManuales.Count-1 do
                            begin
                              qrySentencias_comunicacion.SQL.Clear;
                              qrySentencias_comunicacion.SQL.Add(slSentenciasManuales[i]);
                              qrySentencias_comunicacion.ExecSQL;
                            end;

                            qrySentencias_comunicacion.Connection.CommitTrans;

                            sCadena_error:= chr(13) + 'Factura Insertada y corregida: Empresa:'+sId_empresa+ ' Comprobante: '+ sId_comprobante;
                            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);

                            sCadena_error:= '';
                            if Assigned (slSentencias_comunicacion) then
                              slSentencias_comunicacion.Clear;

                            FreeAndNil(slSentenciasManuales);

                            pSucursal.GraboError           := False;
                            pSucursal.GraboError2Instancia := False;

                            sCadena_error:= '';
                          end;
                        end;
                      except
                        on E: Exception do
                        begin

                          if qrySentencias_comunicacion.Connection.InTransaction then
                            qrySentencias_comunicacion.Connection.RollbackTrans;

                          sCadena_error:= 'No se pudo insertar y corregir la Factura: Empresa:'+sId_empresa+ ' Comprobante: '+ sId_comprobante + Chr(13) + '('+IntToStr(E.HelpContext)+')'+(E.Message)+ Chr(13)+'SQL: ';
                          if Assigned(slSentenciasManuales) then
                          begin
                            for i:=0 to slSentenciasManuales.Count-1 do
                            begin
                              sCadena_error:= sCadena_error + slSentenciasManuales[i] +Chr(13);
                            end;
                            FreeAndNil(slSentenciasManuales);
                          end;
                          if pSucursal.GraboError2Instancia = False then
                          begin
                            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
                            pSucursal.GraboError2Instancia:= true;
                            _CargarItemsMenu;
                          end;
                        end;
                      end;
                      qrySentencias_comunicacion.SQL.Clear;
                    end
                    else
                    begin
                      if pSucursal.GraboError2Instancia = False then
                      begin
                        sCadena_error:= 'No se pudo insertar y corregir la Factura: Empresa:'+sId_empresa+ chr(13)+ ' Comprobante: '+ sId_comprobante;
                        _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
                        pSucursal.GraboError2Instancia:= true;
                        _CargarItemsMenu;
                      end;
                    end;
                  end;
                end

                /////////////////////////////////////////////////
                // Inserta comprobante Faltante
                // Nota de Credito - Referenciada por comprobante de compras - Notas de Recupero de OS
                /////////////////////////////////////////////////
                else if ((PosEx('3---Sucursal:', sCadena_error)> 0)
                  and (PosEx('Cannot add or update a child row: a foreign key constraint fails', sCadena_error)> 0)
                  and (PosEx('`comprobantes_compras`, CONSTRAINT `comprobantes_compras_ibfk_5` FOREIGN KEY (`_ventas_id_empresa`, `_ventas_id_comprobante`) REFERENCES `comprobantes_de_ventas`', sCadena_error)> 0)) then
                begin
                  if slSentencias_comunicacion.Count > 0 then
                  begin

                    // busco la sentencia que contiene el ingreso

                    sId_empresa:='';
                    sId_comprobante:='';
                    for i := 0 to slSentencias_comunicacion.Count-1 do
                    begin
                      if LeftStr(slSentencias_comunicacion[i],27)= 'UPDATE comprobantes_compras' then
                      begin
                        sSentencia:= slSentencias_comunicacion[i];

                        sSentencia:= AnsiReplaceStr(sSentencia,'_latin1','');
                        //sSentencia:= AnsiReplaceStr(sSentencia,Chr(13),'');

                        iPosicion_caracter_inicial:=PosEx('comprobantes_compras._ventas_id_empresa =', sSentencia);
                        if iPosicion_caracter_inicial> -1 then
                        begin
                          iPosicion_caracter_inicial := iPosicion_caracter_inicial + 41;
                          sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, length(sSentencia)-iPosicion_caracter_inicial);
                          iPosicion_caracter_final:=PosEx(',',sSentencia);
                          if iPosicion_caracter_final> -1 then
                          begin
                            sId_empresa:= Trim(copy(sSentencia,0, (iPosicion_caracter_final-1)));
                          end;
                        end;
                        iPosicion_caracter_inicial:=PosEx('comprobantes_compras._ventas_id_comprobante =', sSentencia);
                        if iPosicion_caracter_inicial>-1 then
                        begin
                          iPosicion_caracter_inicial:= iPosicion_caracter_inicial + 47;
                          sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, length(sSentencia)-iPosicion_caracter_inicial);
                          iPosicion_caracter_final:=PosEx(Chr(39),sSentencia);
                          if iPosicion_caracter_final> -1 then
                          begin
                            sId_comprobante:= Trim(copy(sSentencia,1, (iPosicion_caracter_final-1)));
                          end;
                        end;

                        Break;
                      end;
                    end;

                    if ((Length(sId_comprobante)>0) and (Length(sId_empresa)>0)) then
                    begin
                      try

                        qrySentencias_comunicacion.SQL.Clear;

                        slSentenciasManuales:= _Confecciona_Sentencias_Faltantes('nota de credito', sId_empresa, sId_comprobante, '', '', pSucursal);
                        if Assigned(slSentenciasManuales) then
                        begin
                          if slSentenciasManuales.Count>0 then
                          begin
                            qrySentencias_comunicacion.Connection.BeginTrans;
                            for i:=0 to slSentenciasManuales.Count-1 do
                            begin
                              qrySentencias_comunicacion.SQL.Clear;
                              qrySentencias_comunicacion.SQL.Add(slSentenciasManuales[i]);
                              qrySentencias_comunicacion.ExecSQL;
                            end;

                            qrySentencias_comunicacion.Connection.CommitTrans;

                            sCadena_error:= chr(13) + 'Nota de Credito Insertada y corregida: Empresa:'+sId_empresa+ ' Comprobante: '+ sId_comprobante;
                            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);

                            sCadena_error:= '';
                            if Assigned (slSentencias_comunicacion) then
                              slSentencias_comunicacion.Clear;

                            FreeAndNil(slSentenciasManuales);

                            pSucursal.GraboError           := False;
                            pSucursal.GraboError2Instancia := False;

                            sCadena_error:= '';
                          end;
                        end;
                      except
                        on E: Exception do
                        begin

                          if qrySentencias_comunicacion.Connection.InTransaction then
                            qrySentencias_comunicacion.Connection.RollbackTrans;

                          sCadena_error:= 'No se pudo insertar y corregir la Nota de Credito: Empresa:'+sId_empresa+ ' Comprobante: '+ sId_comprobante + Chr(13) + '('+IntToStr(E.HelpContext)+')'+(E.Message)+ Chr(13)+'SQL: ';
                          if Assigned(slSentenciasManuales) then
                          begin
                            for i:=0 to slSentenciasManuales.Count-1 do
                            begin
                              sCadena_error:= sCadena_error + slSentenciasManuales[i] +Chr(13);
                            end;
                            FreeAndNil(slSentenciasManuales);
                          end;
                          if pSucursal.GraboError2Instancia = False then
                          begin
                            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
                            pSucursal.GraboError2Instancia:= true;
                            _CargarItemsMenu;
                          end;
                        end;
                      end;
                      qrySentencias_comunicacion.SQL.Clear;
                    end
                    else
                    begin
                      if pSucursal.GraboError2Instancia = False then
                      begin
                        sCadena_error:= 'No se pudo insertar y corregir la Nota de Credito: Empresa:'+sId_empresa+ chr(13)+ ' Comprobante: '+ sId_comprobante;
                        _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
                        pSucursal.GraboError2Instancia:= true;
                        _CargarItemsMenu;
                      end;
                    end;
                  end;
                end

                /////////////////////////////////////////////////
                // Inserta comprobante Faltante
                // Factura - Referenciada por NC Comprobante para reintegro
                /////////////////////////////////////////////////
                else if ((PosEx('3---Sucursal:', sCadena_error)> 0)
                  and (PosEx('Cannot add or update a child row: a foreign key constraint fails', sCadena_error)> 0)
                  and (PosEx('`notas_credito`, CONSTRAINT `notas_credito_ibfk_3` FOREIGN KEY (`id_empresa`, `_reintegro_id_comprobante`) REFERENCES `comprobantes_de_ventas` (`id_empresa`, `id_comprobante`)', sCadena_error)> 0)) then
                begin
                  if slSentencias_comunicacion.Count > 0 then
                  begin

                    // busco la sentencia que contiene el ingreso

                    sId_empresa:='';
                    sId_comprobante:='';
                    for i := 0 to slSentencias_comunicacion.Count-1 do
                    begin
                      if LeftStr(slSentencias_comunicacion[i],105)= 'INSERT INTO notas_credito (id_empresa, _nc_id_comprobante, _fac_id_comprobante, _reintegro_id_comprobante' then
                      begin
                        sSentencia:= slSentencias_comunicacion[i];
                        iPosicion_caracter_inicial:=PosEx('VALUES (', sSentencia);
                        if iPosicion_caracter_inicial> -1 then
                        begin
                          iPosicion_caracter_inicial := iPosicion_caracter_inicial + 8;
                          sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, length(sSentencia)-iPosicion_caracter_inicial);
                          iPosicion_caracter_final:=PosEx(',',sSentencia);
                          if iPosicion_caracter_final> -1 then
                          begin
                            sId_empresa:= Trim(copy(sSentencia,0, (iPosicion_caracter_final-1)));
                          end;
                        end;
                        // Primer '
                        iPosicion_caracter_inicial:=PosEx(Chr(39), sSentencia);
                        if iPosicion_caracter_inicial>-1 then
                        begin
                          sSentencia:= copy(sSentencia, iPosicion_caracter_inicial+1, length(sSentencia)-iPosicion_caracter_inicial-1);
                          // Segundo '
                          iPosicion_caracter_inicial:=PosEx(Chr(39), sSentencia);
                          if iPosicion_caracter_inicial>-1 then
                          begin
                            sSentencia:= copy(sSentencia, iPosicion_caracter_inicial+1, length(sSentencia)-iPosicion_caracter_inicial-1);
                            // Tercer '
                            iPosicion_caracter_inicial:=PosEx(Chr(39), sSentencia);
                            if iPosicion_caracter_inicial>-1 then
                            begin
                              sSentencia:= copy(sSentencia, iPosicion_caracter_inicial+1, length(sSentencia)-iPosicion_caracter_inicial-1);
                              // Cuarto '
                              iPosicion_caracter_inicial:=PosEx(Chr(39), sSentencia);
                              if iPosicion_caracter_inicial>-1 then
                              begin
                                sSentencia:= copy(sSentencia, iPosicion_caracter_inicial+1, length(sSentencia)-iPosicion_caracter_inicial-1);
                                // Quinto '
                                iPosicion_caracter_inicial:=PosEx(Chr(39), sSentencia);
                                if iPosicion_caracter_inicial>-1 then
                                begin
                                  sSentencia:= copy(sSentencia, iPosicion_caracter_inicial+1, length(sSentencia)-iPosicion_caracter_inicial-1);
                                  iPosicion_caracter_inicial:= 0;
                                  iPosicion_caracter_final:=PosEx(Chr(39),sSentencia);
                                  if iPosicion_caracter_final> -1 then
                                  begin
                                    sId_comprobante:= Trim(copy(sSentencia,0, (iPosicion_caracter_final-1)));
                                  end;
                                end;
                              end;
                            end;
                          end;
                        end;
                        Break;
                      end;
                    end;

                    if ((Length(sId_comprobante)>0) and (Length(sId_empresa)>0)) then
                    begin
                      try

                        qrySentencias_comunicacion.SQL.Clear;

                        slSentenciasManuales:= _Confecciona_Sentencias_Faltantes('factura', sId_empresa, sId_comprobante, '', '', pSucursal);
                        if Assigned(slSentenciasManuales) then
                        begin
                          if slSentenciasManuales.Count>0 then
                          begin
                            qrySentencias_comunicacion.Connection.BeginTrans;
                            for i:=0 to slSentenciasManuales.Count-1 do
                            begin
                              qrySentencias_comunicacion.SQL.Clear;
                              qrySentencias_comunicacion.SQL.Add(slSentenciasManuales[i]);
                              qrySentencias_comunicacion.ExecSQL;
                            end;

                            qrySentencias_comunicacion.Connection.CommitTrans;

                            sCadena_error:= chr(13) + 'Factura Insertada y corregida: Empresa:'+sId_empresa+ ' Comprobante: '+ sId_comprobante;
                            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);

                            sCadena_error:= '';
                            if Assigned (slSentencias_comunicacion) then
                              slSentencias_comunicacion.Clear;

                            FreeAndNil(slSentenciasManuales);

                            pSucursal.GraboError           := False;
                            pSucursal.GraboError2Instancia := False;

                            sCadena_error:= '';
                          end;
                        end;
                      except
                        on E: Exception do
                        begin

                          if qrySentencias_comunicacion.Connection.InTransaction then
                            qrySentencias_comunicacion.Connection.RollbackTrans;

                          sCadena_error:= 'No se pudo insertar y corregir la Factura: Empresa:'+sId_empresa+ ' Comprobante: '+ sId_comprobante + Chr(13) + '('+IntToStr(E.HelpContext)+')'+(E.Message)+ Chr(13)+'SQL: ';
                          if Assigned(slSentenciasManuales) then
                          begin
                            for i:=0 to slSentenciasManuales.Count-1 do
                            begin
                              sCadena_error:= sCadena_error + slSentenciasManuales[i] +Chr(13);
                            end;
                            FreeAndNil(slSentenciasManuales);
                          end;
                          if pSucursal.GraboError2Instancia = False then
                          begin
                            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
                            pSucursal.GraboError2Instancia:= true;
                            _CargarItemsMenu;
                          end;
                        end;
                      end;
                      qrySentencias_comunicacion.SQL.Clear;
                    end
                    else
                    begin
                      if pSucursal.GraboError2Instancia = False then
                      begin
                        sCadena_error:= 'No se pudo insertar y corregir la Factura: Empresa:'+sId_empresa+ chr(13)+ ' Comprobante: '+ sId_comprobante;
                        _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
                        pSucursal.GraboError2Instancia:= true;
                        _CargarItemsMenu;
                      end;
                    end;
                  end;
                end
                /////////////////////////////////////////////////
                // Inserta comprobante Faltante
                // Factura - Referenciada por NC Comprobante para reintegro
                /////////////////////////////////////////////////
                else if ((PosEx('3---Sucursal:', sCadena_error)> 0)
                  and (PosEx('Cannot add or update a child row: a foreign key constraint fails', sCadena_error)> 0)
                  and (PosEx('`comprobantes_vta_mov_sdo_favor`, CONSTRAINT `comprobantes_vta_mov_sdo_favor_fk1` FOREIGN KEY (`id_empresa`, `_origen_id_comprobante`) REFERENCES `comprobantes_de_ventas`', sCadena_error)> 0)) then
                begin
                  if slSentencias_comunicacion.Count > 0 then
                  begin

                    // busco la sentencia que contiene el comprobante

                    sId_empresa:='';
                    sId_comprobante:='';
                    for i := 0 to slSentencias_comunicacion.Count-1 do
                    begin
                      if LeftStr(slSentencias_comunicacion[i],217)= 'INSERT INTO comprobantes_vta_mov_sdo_favor (id_cupon, id_empresa, _origen_id_comprobante, _origen_vale_cf_id_empresa, _origen_vale_cf_id_comprobante, _destino_id_empresa, _destino_id_comprobante, importe_pago, estado)' then
                      begin
                        sSentencia:= slSentencias_comunicacion[i];
                        iPosicion_caracter_inicial:=PosEx('VALUES (', sSentencia);
                        if iPosicion_caracter_inicial> -1 then
                        begin
                          iPosicion_caracter_inicial := iPosicion_caracter_inicial + 8;
                          sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, length(sSentencia)-iPosicion_caracter_inicial);
                        end;

                        // Primera ,
                        iPosicion_caracter_inicial:=PosEx(',', sSentencia);
                        if iPosicion_caracter_inicial>-1 then
                        begin
                          sSentencia:= copy(sSentencia, iPosicion_caracter_inicial+1, length(sSentencia)-iPosicion_caracter_inicial-1);
                          // Segunda ,
                          iPosicion_caracter_inicial:=PosEx(',', sSentencia);
                          if iPosicion_caracter_inicial>-1 then
                          begin
                            sSentencia:= copy(sSentencia, iPosicion_caracter_inicial+1, length(sSentencia)-iPosicion_caracter_inicial-1);

                            iPosicion_caracter_inicial:= 0;
                            iPosicion_caracter_final:=PosEx(',',sSentencia);
                            if iPosicion_caracter_final> -1 then
                            begin
                              sId_empresa:= Trim(copy(sSentencia,0, (iPosicion_caracter_final-1)));
                            end;

                            // Sexta ,
                            iPosicion_caracter_inicial:=PosEx(',', sSentencia);
                            if iPosicion_caracter_inicial>-1 then
                            begin
                              sSentencia:= copy(sSentencia, iPosicion_caracter_inicial+1, length(sSentencia)-iPosicion_caracter_inicial-1);
                              // Primer '
                              iPosicion_caracter_inicial:=PosEx(Chr(39), sSentencia);
                              if iPosicion_caracter_inicial>-1 then
                              begin
                                sSentencia:= copy(sSentencia, iPosicion_caracter_inicial+1, length(sSentencia)-iPosicion_caracter_inicial-1);

                                iPosicion_caracter_inicial:= 0;
                                iPosicion_caracter_final:=PosEx(Chr(39),sSentencia);
                                if iPosicion_caracter_final> -1 then
                                begin
                                  sId_comprobante:= Trim(copy(sSentencia,0, (iPosicion_caracter_final-1)));
                                end;
                              end;
                            end;
                          end;
                        end;

                        Break;
                      end;
                    end;

                    if ((Length(Trim(sId_comprobante))>0) and (Length(Trim(sId_empresa))>0)) then
                    begin
                      try

                        qrySentencias_comunicacion.SQL.Clear;

                        slSentenciasManuales:= _Confecciona_Sentencias_Faltantes('factura', sId_empresa, sId_comprobante, '', '', pSucursal);
                        if Assigned(slSentenciasManuales) then
                        begin
                          if slSentenciasManuales.Count>0 then
                          begin
                            qrySentencias_comunicacion.Connection.BeginTrans;
                            for i:=0 to slSentenciasManuales.Count-1 do
                            begin
                              qrySentencias_comunicacion.SQL.Clear;
                              qrySentencias_comunicacion.SQL.Add(slSentenciasManuales[i]);
                              qrySentencias_comunicacion.ExecSQL;
                            end;

                            qrySentencias_comunicacion.Connection.CommitTrans;

                            sCadena_error:= chr(13) + 'Factura Insertada y corregida: Empresa:'+sId_empresa+ ' Comprobante: '+ sId_comprobante;
                            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);

                            sCadena_error:= '';
                            if Assigned (slSentencias_comunicacion) then
                              slSentencias_comunicacion.Clear;

                            FreeAndNil(slSentenciasManuales);

                            pSucursal.GraboError           := False;
                            pSucursal.GraboError2Instancia := False;

                            sCadena_error:= '';
                          end;
                        end;
                      except
                        on E: Exception do
                        begin

                          if qrySentencias_comunicacion.Connection.InTransaction then
                            qrySentencias_comunicacion.Connection.RollbackTrans;

                          sCadena_error:= 'No se pudo insertar y corregir la Factura: Empresa:'+sId_empresa+ ' Comprobante: '+ sId_comprobante + Chr(13) + '('+IntToStr(E.HelpContext)+')'+(E.Message)+ Chr(13)+'SQL: ';
                          if Assigned(slSentenciasManuales) then
                          begin
                            for i:=0 to slSentenciasManuales.Count-1 do
                            begin
                              sCadena_error:= sCadena_error + slSentenciasManuales[i] +Chr(13);
                            end;
                            FreeAndNil(slSentenciasManuales);
                          end;
                          if pSucursal.GraboError2Instancia = False then
                          begin
                            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
                            pSucursal.GraboError2Instancia:= true;
                            _CargarItemsMenu;
                          end;
                        end;
                      end;
                      qrySentencias_comunicacion.SQL.Clear;
                    end
                    else
                    begin
                      if pSucursal.GraboError2Instancia = False then
                      begin
                        sCadena_error:= 'No se pudo insertar y corregir la Factura: Empresa:'+sId_empresa+ chr(13)+ ' Comprobante: '+ sId_comprobante;
                        _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
                        pSucursal.GraboError2Instancia:= true;
                        _CargarItemsMenu;
                      end;
                    end;
                  end;
                end
                /////////////////////////////////////////////////
                // Inserta comprobante Faltante
                // Factura - Referenciada por Vale de Entrega
                /////////////////////////////////////////////////
                else if ((PosEx('3---Sucursal:', sCadena_error)> 0)
                  and (PosEx('Cannot add or update a child row: a foreign key constraint fails', sCadena_error)> 0)
                  and (PosEx('`movimientos_stock`, CONSTRAINT `movimientos_stock_ibfk_6` FOREIGN KEY (`_vales_id_comprobante`, `_vales_item`) REFERENCES `vales_de_entrega`', sCadena_error)> 0)) then
                begin
                  if slSentencias_comunicacion.Count > 0 then
                  begin

                    // busco la sentencia que contiene el ingreso
                    sId_empresa:='';
                    sId_comprobante:='';
                    for i := 0 to slSentencias_comunicacion.Count-1 do
                    begin
                      if LeftStr(slSentencias_comunicacion[i],23)= 'UPDATE vales_de_entrega' then
                      begin
                        sSentencia:= slSentencias_comunicacion[i];
                        iPosicion_caracter_inicial:=PosEx('id_empresa =', sSentencia);
                        if iPosicion_caracter_inicial> -1 then
                        begin
                          iPosicion_caracter_inicial := iPosicion_caracter_inicial + 12;
                          sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, length(sSentencia)-iPosicion_caracter_inicial);
                          iPosicion_caracter_final:=PosEx('AND',sSentencia);
                          if iPosicion_caracter_final> -1 then
                          begin
                            sId_empresa:= Trim(copy(sSentencia,0, (iPosicion_caracter_final-1)));
                          end;
                        end;
                        // Primer '
                        iPosicion_caracter_inicial:=PosEx(Chr(39), sSentencia);
                        if iPosicion_caracter_inicial>-1 then
                        begin
                          sSentencia:= copy(sSentencia, iPosicion_caracter_inicial+1, length(sSentencia)-iPosicion_caracter_inicial-1);
                          iPosicion_caracter_inicial:= 0;
                          iPosicion_caracter_final:=PosEx(Chr(39),sSentencia);
                          if iPosicion_caracter_final> -1 then
                          begin
                            sId_comprobante:= Trim(copy(sSentencia,0, (iPosicion_caracter_final-1)));
                          end;
                        end;
                        Break;
                      end;
                    end;

                    if ((Length(sId_comprobante)>0) and (Length(sId_empresa)>0)) then
                    begin
                      try

                        qrySentencias_comunicacion.SQL.Clear;

                        slSentenciasManuales:= _Confecciona_Sentencias_Faltantes('factura', sId_empresa, sId_comprobante, '', '', pSucursal);
                        if Assigned(slSentenciasManuales) then
                        begin
                          if slSentenciasManuales.Count>0 then
                          begin
                            qrySentencias_comunicacion.Connection.BeginTrans;
                            for i:=0 to slSentenciasManuales.Count-1 do
                            begin
                              qrySentencias_comunicacion.SQL.Clear;
                              qrySentencias_comunicacion.SQL.Add(slSentenciasManuales[i]);
                              qrySentencias_comunicacion.ExecSQL;
                            end;

                            qrySentencias_comunicacion.Connection.CommitTrans;

                            sCadena_error:= chr(13) + 'Factura Insertada y corregida: Empresa:'+sId_empresa+ ' Comprobante: '+ sId_comprobante;
                            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);

                            sCadena_error:= '';
                            if Assigned (slSentencias_comunicacion) then
                              slSentencias_comunicacion.Clear;

                            FreeAndNil(slSentenciasManuales);

                            pSucursal.GraboError           := False;
                            pSucursal.GraboError2Instancia := False;

                            sCadena_error:= '';
                          end;
                        end;
                      except
                        on E: Exception do
                        begin

                          if qrySentencias_comunicacion.Connection.InTransaction then
                            qrySentencias_comunicacion.Connection.RollbackTrans;

                          sCadena_error:= 'No se pudo insertar y corregir la Factura: Empresa:'+sId_empresa+ ' Comprobante: '+ sId_comprobante + Chr(13) + '('+IntToStr(E.HelpContext)+')'+(E.Message)+ Chr(13)+'SQL: ';
                          if Assigned(slSentenciasManuales) then
                          begin
                            for i:=0 to slSentenciasManuales.Count-1 do
                            begin
                              sCadena_error:= sCadena_error + slSentenciasManuales[i] +Chr(13);
                            end;
                            FreeAndNil(slSentenciasManuales);
                          end;
                          if pSucursal.GraboError2Instancia = False then
                          begin
                            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
                            pSucursal.GraboError2Instancia:= true;
                            _CargarItemsMenu;
                          end;
                        end;
                      end;
                      qrySentencias_comunicacion.SQL.Clear;
                    end
                    else
                    begin
                      if pSucursal.GraboError2Instancia = False then
                      begin
                        sCadena_error:= 'No se pudo insertar y corregir la Factura: Empresa:'+sId_empresa+ chr(13)+ ' Comprobante: '+ sId_comprobante;
                        _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
                        pSucursal.GraboError2Instancia:= true;
                        _CargarItemsMenu;
                      end;
                    end;
                  end;
                end

                /////////////////////////////////////////////////
                // Inserta comprobante de compras Faltante referenciado como comision bancaria
                // Comprobante de Compras - Referenciada por orden de pago - Comisiones
                /////////////////////////////////////////////////
                else if ((PosEx('3---Sucursal:', sCadena_error)> 0)
                  and (PosEx('Cannot add or update a child row: a foreign key constraint fails', sCadena_error)> 0)
                  and (PosEx('`comisiones_bancarias`, CONSTRAINT `comisiones_bancarias_fk2` FOREIGN KEY (`id_proveedor`, `id_comprobante`) REFERENCES `comprobantes_compras`', sCadena_error)> 0)) then
                begin
                  if slSentencias_comunicacion.Count > 0 then
                  begin

                    // busco la sentencia que contiene la comision
                    sId_proveedor:='';
                    sId_comprobante:='';
                    for i := 0 to slSentencias_comunicacion.Count-1 do
                    begin
                      if LeftStr(slSentencias_comunicacion[i],32)= 'INSERT INTO comisiones_bancarias' then
                      begin
                        sSentencia:= slSentencias_comunicacion[i];
                        sSentencia:= AnsiReplaceStr(sSentencia,'_latin1','');

                        iPosicion_caracter_inicial:=PosEx('VALUES (', sSentencia);
                        if iPosicion_caracter_inicial> -1 then
                        begin
                          iPosicion_caracter_inicial := iPosicion_caracter_inicial + 8;
                          sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, length(sSentencia)-iPosicion_caracter_inicial+1);

                          // Primera ,
                          iPosicion_caracter_inicial:=PosEx(',', sSentencia);
                          if iPosicion_caracter_inicial >= -1 then
                          begin
                            iPosicion_caracter_inicial:= iPosicion_caracter_inicial + 1;
                            sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, length(sSentencia)-iPosicion_caracter_inicial+1);

                            // Segunda ,
                            iPosicion_caracter_inicial:=PosEx(',', sSentencia);
                            if iPosicion_caracter_inicial >= -1 then
                            begin
                              iPosicion_caracter_inicial:= iPosicion_caracter_inicial + 1;
                              sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, length(sSentencia)-iPosicion_caracter_inicial+1);

                              iPosicion_caracter_final:=PosEx(',',sSentencia);
                              if iPosicion_caracter_final> -1 then
                              begin
                                sId_proveedor:= Trim(copy(sSentencia,0, (iPosicion_caracter_final-1)));
                              end;

                              // Tercera ,
                              iPosicion_caracter_inicial:=PosEx(',', sSentencia);
                              if iPosicion_caracter_inicial >= -1 then
                              begin
                                iPosicion_caracter_inicial:= iPosicion_caracter_inicial + 1;
                                sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, length(sSentencia)-iPosicion_caracter_inicial+1);

                                iPosicion_caracter_final:=PosEx(',',sSentencia);
                                if iPosicion_caracter_final> -1 then
                                begin
                                  sId_comprobante:= Trim(copy(sSentencia,0, (iPosicion_caracter_final-1)));
                                  sId_comprobante:= AnsiReplaceStr(sId_comprobante,Chr(39),'');
                                end;
                              end;
                            end;
                          end;
                        end;

                        Break;
                      end;
                    end;

                    if ((Length(sId_comprobante)>0) and (Length(sId_proveedor)>0)) then
                    begin
                      try

                        qrySentencias_comunicacion.SQL.Clear;

                        slSentenciasManuales:= _Confecciona_Sentencias_Faltantes('comprobante_compra', sId_proveedor, sId_comprobante, '', '', pSucursal);
                        if Assigned(slSentenciasManuales) then
                        begin
                          if slSentenciasManuales.Count>0 then
                          begin
                            qrySentencias_comunicacion.Connection.BeginTrans;
                            for i:=0 to slSentenciasManuales.Count-1 do
                            begin
                              qrySentencias_comunicacion.SQL.Clear;
                              qrySentencias_comunicacion.SQL.Add(slSentenciasManuales[i]);
                              qrySentencias_comunicacion.ExecSQL;
                            end;

                            qrySentencias_comunicacion.Connection.CommitTrans;

                            sCadena_error:= chr(13) + 'Comprobante de compras Insertado y corregida: Proveedor:'+sId_proveedor+ ' Comprobante: '+ sId_comprobante;
                            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);

                            sCadena_error:= '';
                            if Assigned (slSentencias_comunicacion) then
                              slSentencias_comunicacion.Clear;

                            FreeAndNil(slSentenciasManuales);

                            pSucursal.GraboError           := False;
                            pSucursal.GraboError2Instancia := False;

                            sCadena_error:= '';
                          end;
                        end;
                      except
                        on E: Exception do
                        begin

                          if qrySentencias_comunicacion.Connection.InTransaction then
                            qrySentencias_comunicacion.Connection.RollbackTrans;

                          sCadena_error:= 'No se pudo insertar y corregir el Comprobante de compra: Proveedor:'+sId_proveedor+ ' Comprobante: '+ sId_comprobante + Chr(13) + '('+IntToStr(E.HelpContext)+')'+(E.Message)+ Chr(13)+'SQL: ';
                          if Assigned(slSentenciasManuales) then
                          begin
                            for i:=0 to slSentenciasManuales.Count-1 do
                            begin
                              sCadena_error:= sCadena_error + slSentenciasManuales[i] +Chr(13);
                            end;
                            FreeAndNil(slSentenciasManuales);
                          end;
                          if pSucursal.GraboError2Instancia = False then
                          begin
                            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
                            pSucursal.GraboError2Instancia:= true;
                            _CargarItemsMenu;
                          end;
                        end;
                      end;
                      qrySentencias_comunicacion.SQL.Clear;
                    end
                    else
                    begin
                      if pSucursal.GraboError2Instancia = False then
                      begin
                        sCadena_error:= 'No se pudo insertar y corregir el Comprobante de compra: Proveedor:'+sId_proveedor+ chr(13)+ ' Comprobante: '+ sId_comprobante;
                        _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
                        pSucursal.GraboError2Instancia:= true;
                        _CargarItemsMenu;
                      end;
                    end;
                  end;
                end

                /////////////////////////////////////////////////
                // Inserta comprobante de compras Faltante referenciado como comprobantes_compras_por_nc
                // Comprobante de Compras - Referenciada por orden de pago - comprobantes_compras_por_nc
                /////////////////////////////////////////////////
                else if ((PosEx('3---Sucursal:', sCadena_error)> 0)
                  and (PosEx('Cannot add or update a child row: a foreign key constraint fails', sCadena_error)> 0)
                  and (PosEx('`comprobantes_compras_por_nc`, CONSTRAINT `comprobantes_compras_por_nc_fk1` FOREIGN KEY (`_nc_id_proveedor`, `_nc_id_comprobante`) REFERENCES `comprobantes_compras`', sCadena_error)> 0)) then
                begin
                  if slSentencias_comunicacion.Count > 0 then
                  begin

                    // busco la sentencia que contiene la comision
                    sId_proveedor:='';
                    sId_comprobante:='';
                    for i := 0 to slSentencias_comunicacion.Count-1 do
                    begin
                      if ((LeftStr(slSentencias_comunicacion[i],11)= 'INSERT INTO') and (PosEx('comprobantes_compras_por_nc', slSentencias_comunicacion[i])> 0)) then
                      begin

                        sId_proveedor:='';
                        sId_comprobante:='';

                        sSentencia:= slSentencias_comunicacion[i];
                        sSentencia:= AnsiReplaceStr(sSentencia,'_latin1','');

                        iPosicion_caracter_inicial:=PosEx('VALUES (', sSentencia);
                        if iPosicion_caracter_inicial> -1 then
                        begin
                          iPosicion_caracter_inicial := iPosicion_caracter_inicial + 8;
                          sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, length(sSentencia)-iPosicion_caracter_inicial+1);

                          // Primera ,
                          iPosicion_caracter_inicial:=PosEx(',', sSentencia);
                          if iPosicion_caracter_inicial >= -1 then
                          begin
                            iPosicion_caracter_inicial:= iPosicion_caracter_inicial + 1;
                            sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, length(sSentencia)-iPosicion_caracter_inicial+1);

                            // Segunda ,
                            iPosicion_caracter_inicial:=PosEx(',', sSentencia);
                            if iPosicion_caracter_inicial >= -1 then
                            begin
                              iPosicion_caracter_inicial:= iPosicion_caracter_inicial + 1;
                              sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, length(sSentencia)-iPosicion_caracter_inicial+1);

                              iPosicion_caracter_final:=PosEx(',',sSentencia);
                              if iPosicion_caracter_final> -1 then
                              begin
                                sId_proveedor:= Trim(copy(sSentencia,0, (iPosicion_caracter_final-1)));
                              end;

                              // Tercera ,
                              iPosicion_caracter_inicial:=PosEx(',', sSentencia);
                              if iPosicion_caracter_inicial >= -1 then
                              begin
                                iPosicion_caracter_inicial:= iPosicion_caracter_inicial + 1;
                                sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, length(sSentencia)-iPosicion_caracter_inicial+1);

                                iPosicion_caracter_final:=PosEx(',',sSentencia);
                                if iPosicion_caracter_final> -1 then
                                begin
                                  sId_comprobante:= Trim(copy(sSentencia,0, (iPosicion_caracter_final-1)));
                                  sId_comprobante:= AnsiReplaceStr(sId_comprobante,Chr(39),'');
                                end;
                              end;
                            end;
                          end;
                        end;

                        if ((Length(sId_comprobante)>0) and (Length(sId_proveedor)>0)) then
                        begin
                          try

                            qrySentencias_comunicacion.SQL.Clear;

                            slSentenciasManuales:= _Confecciona_Sentencias_Faltantes('comprobante_compra', sId_proveedor, sId_comprobante, '', '', pSucursal);
                            if Assigned(slSentenciasManuales) then
                            begin
                              if slSentenciasManuales.Count>0 then
                              begin
                                qrySentencias_comunicacion.Connection.BeginTrans;
                                for j:=0 to slSentenciasManuales.Count-1 do
                                begin
                                  qrySentencias_comunicacion.SQL.Clear;
                                  qrySentencias_comunicacion.SQL.Add(slSentenciasManuales[j]);
                                  qrySentencias_comunicacion.ExecSQL;
                                end;

                                qrySentencias_comunicacion.Connection.CommitTrans;

                                sCadena_error:= chr(13) + 'Comprobante de compras Insertado y corregida: Proveedor:'+sId_proveedor+ ' Comprobante: '+ sId_comprobante;
                                _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);

                                sCadena_error:= '';
                                //if Assigned (slSentencias_comunicacion) then
                                //  slSentencias_comunicacion.Clear;

                                FreeAndNil(slSentenciasManuales);

                                pSucursal.GraboError           := False;
                                pSucursal.GraboError2Instancia := False;

                                sCadena_error:= '';
                              end;
                            end;
                          except
                            on E: Exception do
                            begin

                              if qrySentencias_comunicacion.Connection.InTransaction then
                                qrySentencias_comunicacion.Connection.RollbackTrans;

                              sCadena_error:= 'No se pudo insertar y corregir el Comprobante de compra: Proveedor:'+sId_proveedor+ ' Comprobante: '+ sId_comprobante + Chr(13) + '('+IntToStr(E.HelpContext)+')'+(E.Message)+ Chr(13)+'SQL: ';
                              if Assigned(slSentenciasManuales) then
                              begin
                                for j:=0 to slSentenciasManuales.Count-1 do
                                begin
                                  sCadena_error:= sCadena_error + slSentenciasManuales[j] +Chr(13);
                                end;
                                FreeAndNil(slSentenciasManuales);
                              end;
                              if pSucursal.GraboError2Instancia = False then
                              begin
                                _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
                                pSucursal.GraboError2Instancia:= False;
                                //_CargarItemsMenu;
                              end;
                            end;
                          end;
                          qrySentencias_comunicacion.SQL.Clear;
                        end
                        else
                        begin
                          if pSucursal.GraboError2Instancia = False then
                          begin
                            sCadena_error:= 'No se pudo insertar y corregir el Comprobante de compra: Proveedor:'+sId_proveedor+ chr(13)+ ' Comprobante: '+ sId_comprobante;
                            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
                            pSucursal.GraboError2Instancia:= False;
                            //_CargarItemsMenu;
                          end;
                        end;

                      end;
                    end;

                    if Assigned (slSentencias_comunicacion) then
                      slSentencias_comunicacion.Clear;

                  end;
                end


                /////////////////////////////////////////////////
                // Inserta comprobante de compras Faltante referenciado por lin_orden_pago_a_proveedores
                // Comprobante de Compras - Referenciada por orden de pago - lin_orden_pago_a_proveedores
                /////////////////////////////////////////////////
                else if ((PosEx('3---Sucursal:', sCadena_error)> 0)
                  and (PosEx('Cannot add or update a child row: a foreign key constraint fails', sCadena_error)> 0)
                  and (PosEx('`lin_orden_pago_a_proveedores`, CONSTRAINT `lin_orden_pago_a_proveedores_ibfk_1` FOREIGN KEY (`id_proveedor`, `id_comprobante`) REFERENCES `comprobantes_compras`', sCadena_error)> 0)) then
                begin
                  if slSentencias_comunicacion.Count > 0 then
                  begin

                    // busco la sentencia que contiene la comision
                    sId_proveedor:='';
                    sId_comprobante:='';
                    for i := 0 to slSentencias_comunicacion.Count-1 do
                    begin
                      if ((LeftStr(slSentencias_comunicacion[i],11)= 'INSERT INTO') and (PosEx('lin_orden_pago_a_proveedores', slSentencias_comunicacion[i])> 0)) then
                      begin

                        sId_proveedor:='';
                        sId_comprobante:='';

                        sSentencia:= slSentencias_comunicacion[i];
                        sSentencia:= AnsiReplaceStr(sSentencia,'_latin1','');

                        iPosicion_caracter_inicial:=PosEx('VALUES', sSentencia);
                        if iPosicion_caracter_inicial> -1 then
                        begin
                          iPosicion_caracter_inicial := iPosicion_caracter_inicial + 6;
                          sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, length(sSentencia)-iPosicion_caracter_inicial+1);

                          // Primera ,
                          iPosicion_caracter_inicial:=PosEx(',', sSentencia);
                          if iPosicion_caracter_inicial >= -1 then
                          begin
                            iPosicion_caracter_inicial:= iPosicion_caracter_inicial + 1;
                            sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, length(sSentencia)-iPosicion_caracter_inicial+1);

                            // Segunda ,
                            iPosicion_caracter_inicial:=PosEx(',', sSentencia);
                            if iPosicion_caracter_inicial >= -1 then
                            begin
                              iPosicion_caracter_inicial:= iPosicion_caracter_inicial + 1;
                              sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, length(sSentencia)-iPosicion_caracter_inicial+1);

                              // Tercera ,
                              iPosicion_caracter_inicial:=PosEx(',', sSentencia);
                              if iPosicion_caracter_inicial >= -1 then
                              begin
                                iPosicion_caracter_inicial:= iPosicion_caracter_inicial + 1;
                                sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, length(sSentencia)-iPosicion_caracter_inicial+1);

                                iPosicion_caracter_final:=PosEx(',',sSentencia);
                                if iPosicion_caracter_final> -1 then
                                begin
                                  sId_comprobante:= Trim(copy(sSentencia,0, (iPosicion_caracter_final-1)));
                                  sId_comprobante:= AnsiReplaceStr(sId_comprobante,Chr(39),'');
                                end;

                                // Cuarta ,
                                iPosicion_caracter_inicial:=PosEx(',', sSentencia);
                                if iPosicion_caracter_inicial >= -1 then
                                begin
                                  iPosicion_caracter_inicial:= iPosicion_caracter_inicial + 1;
                                  sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, length(sSentencia)-iPosicion_caracter_inicial+1);

                                  iPosicion_caracter_final:=PosEx(',',sSentencia);
                                  if iPosicion_caracter_final> -1 then
                                  begin
                                    sId_proveedor:= Trim(copy(sSentencia,0, (iPosicion_caracter_final-1)));
                                  end;

                                end;
                              end;
                            end;
                          end;
                        end;

                        if ((Length(sId_comprobante)>0) and (Length(sId_proveedor)>0)) then
                        begin
                          try

                            qrySentencias_comunicacion.SQL.Clear;

                            slSentenciasManuales:= _Confecciona_Sentencias_Faltantes('comprobante_compra', sId_proveedor, sId_comprobante, '', '', pSucursal);
                            if Assigned(slSentenciasManuales) then
                            begin
                              if slSentenciasManuales.Count>0 then
                              begin
                                qrySentencias_comunicacion.Connection.BeginTrans;
                                for j:=0 to slSentenciasManuales.Count-1 do
                                begin
                                  qrySentencias_comunicacion.SQL.Clear;
                                  qrySentencias_comunicacion.SQL.Add(slSentenciasManuales[j]);
                                  qrySentencias_comunicacion.ExecSQL;
                                end;

                                qrySentencias_comunicacion.Connection.CommitTrans;

                                sCadena_error:= chr(13) + 'Comprobante de compras Insertado y corregida: Proveedor:'+sId_proveedor+ ' Comprobante: '+ sId_comprobante;
                                _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);

                                sCadena_error:= '';
                                //if Assigned (slSentencias_comunicacion) then
                                //  slSentencias_comunicacion.Clear;

                                FreeAndNil(slSentenciasManuales);

                                pSucursal.GraboError           := False;
                                pSucursal.GraboError2Instancia := False;

                                sCadena_error:= '';
                              end;
                            end;
                          except
                            on E: Exception do
                            begin

                              if qrySentencias_comunicacion.Connection.InTransaction then
                                qrySentencias_comunicacion.Connection.RollbackTrans;

                              sCadena_error:= 'No se pudo insertar y corregir el Comprobante de compra: Proveedor:'+sId_proveedor+ ' Comprobante: '+ sId_comprobante + Chr(13) + '('+IntToStr(E.HelpContext)+')'+(E.Message)+ Chr(13)+'SQL: ';
                              if Assigned(slSentenciasManuales) then
                              begin
                                for j:=0 to slSentenciasManuales.Count-1 do
                                begin
                                  sCadena_error:= sCadena_error + slSentenciasManuales[j] +Chr(13);
                                end;
                                FreeAndNil(slSentenciasManuales);
                              end;
                              if pSucursal.GraboError2Instancia = False then
                              begin
                                _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
                                pSucursal.GraboError2Instancia:= False;
                                //_CargarItemsMenu;
                              end;
                            end;
                          end;
                          qrySentencias_comunicacion.SQL.Clear;
                        end
                        else
                        begin
                          if pSucursal.GraboError2Instancia = False then
                          begin
                            sCadena_error:= 'No se pudo insertar y corregir el Comprobante de compra: Proveedor:'+sId_proveedor+ chr(13)+ ' Comprobante: '+ sId_comprobante;
                            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
                            pSucursal.GraboError2Instancia:= False;
                            //_CargarItemsMenu;
                          end;
                        end;

                      end;
                    end;

                    if Assigned (slSentencias_comunicacion) then
                      slSentencias_comunicacion.Clear;

                  end;
                end


                /////////////////////////////////////////////////
                // Inserta persona referenciada como _fidelizacion_id_persona
                // Persona - Referenciada por comprobantes_de_ventas _fidelizacion_id_persona
                /////////////////////////////////////////////////
                else if ((PosEx('3---Sucursal:', sCadena_error)> 0)
                  and (PosEx('Cannot add or update a child row: a foreign key constraint fails', sCadena_error)> 0)
                  and (PosEx('`comprobantes_de_ventas`, CONSTRAINT `comprobantes_de_ventas_ibfk_6` FOREIGN KEY (`_fidelizacion_id_persona`) REFERENCES `personas`', sCadena_error)> 0)) then
                begin
                  if slSentencias_comunicacion.Count > 0 then
                  begin

                    // busco la sentencia que contiene la persona
                    sId_Persona:='';
                    for i := 0 to slSentencias_comunicacion.Count-1 do
                    begin
                      if LeftStr(slSentencias_comunicacion[i],34)= 'INSERT INTO comprobantes_de_ventas' then
                      begin
                        sSentencia:= slSentencias_comunicacion[i];
                        sSentencia:= AnsiReplaceStr(sSentencia,'_latin1','');

                        iPosicion_caracter_inicial:=PosEx('VALUES (', sSentencia);
                        if iPosicion_caracter_inicial> -1 then
                        begin

                          if LeftStr(sSentencia,82)='INSERT INTO comprobantes_de_ventas (id_empresa,id_comprobante,id_sucursal,nro_caja' then
                            iCantidad:= 41
                          else
                            iCantidad:= 15;

                          iPosicion_caracter_inicial := iPosicion_caracter_inicial + 8;
                          sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, length(sSentencia)-iPosicion_caracter_inicial+1);

                          for j := 1 to iCantidad do
                          begin
                            iPosicion_caracter_inicial:=PosEx(',', sSentencia);
                            if iPosicion_caracter_inicial >= -1 then
                            begin
                              iPosicion_caracter_inicial:= iPosicion_caracter_inicial + 1;
                              sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, length(sSentencia)-iPosicion_caracter_inicial+1);
                            end;
                          end;

                          // Coma final,
                          iPosicion_caracter_final:=PosEx(',',sSentencia);
                          if iPosicion_caracter_final> -1 then
                          begin
                            sId_persona:= Trim(copy(sSentencia,0, (iPosicion_caracter_final-1)));
                          end;

                        end;

                        Break;
                      end;
                    end;

                    if Length(sId_Persona)>0 then
                    begin
                      try

                        qrySentencias_comunicacion.SQL.Clear;

                        slSentenciasManuales:= _Confecciona_Sentencias_Faltantes('persona', sId_persona, '', '', '', pSucursal);
                        if Assigned(slSentenciasManuales) then
                        begin
                          if slSentenciasManuales.Count>0 then
                          begin
                            qrySentencias_comunicacion.Connection.BeginTrans;
                            for i:=0 to slSentenciasManuales.Count-1 do
                            begin
                              qrySentencias_comunicacion.SQL.Clear;
                              qrySentencias_comunicacion.SQL.Add(slSentenciasManuales[i]);
                              qrySentencias_comunicacion.ExecSQL;
                            end;

                            qrySentencias_comunicacion.Connection.CommitTrans;

                            sCadena_error:= chr(13) + 'Persona Registrada y corregida: Persona:'+sId_Persona;
                            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);

                            sCadena_error:= '';
                            if Assigned (slSentencias_comunicacion) then
                              slSentencias_comunicacion.Clear;

                            FreeAndNil(slSentenciasManuales);

                            pSucursal.GraboError           := False;
                            pSucursal.GraboError2Instancia := False;

                            sCadena_error:= '';
                          end;
                        end;
                      except
                        on E: Exception do
                        begin

                          if qrySentencias_comunicacion.Connection.InTransaction then
                            qrySentencias_comunicacion.Connection.RollbackTrans;

                          sCadena_error:= 'No se pudo registrar y corregir la Persona: '+sId_Persona + Chr(13) + '('+IntToStr(E.HelpContext)+')'+(E.Message)+ Chr(13)+'SQL: ';
                          if Assigned(slSentenciasManuales) then
                          begin
                            for i:=0 to slSentenciasManuales.Count-1 do
                            begin
                              sCadena_error:= sCadena_error + slSentenciasManuales[i] +Chr(13);
                            end;
                            FreeAndNil(slSentenciasManuales);
                          end;
                          if pSucursal.GraboError2Instancia = False then
                          begin
                            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
                            pSucursal.GraboError2Instancia:= true;
                            _CargarItemsMenu;
                          end;
                        end;
                      end;
                      qrySentencias_comunicacion.SQL.Clear;
                    end
                    else
                    begin
                      if pSucursal.GraboError2Instancia = False then
                      begin
                        sCadena_error:= 'No se pudo registrar y corregir la Persona: '+sId_Persona;
                        _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
                        pSucursal.GraboError2Instancia:= true;
                        _CargarItemsMenu;
                      end;
                    end;
                  end;
                end

                /////////////////////////////////////////////////
                // Inserta registro de caja referenciado por comprobante de ventas
                // Registro de caja
                /////////////////////////////////////////////////
                else if ((PosEx('3---Sucursal:', sCadena_error)> 0)
                  and (PosEx('Cannot add or update a child row: a foreign key constraint fails', sCadena_error)> 0)
                  and (PosEx('`comprobantes_de_ventas`, CONSTRAINT `comprobantes_de_ventas_ibfk_3` FOREIGN KEY (`id_empresa`, `id_sucursal`, `nro_caja`, `periodo_caja`) REFERENCES `registro_de_cajas`', sCadena_error)> 0)) then
                begin
                  if slSentencias_comunicacion.Count > 0 then
                  begin

                    // busco la sentencia que contiene los datos del registro de caja
                    sId_empresa:='';
                    sId_Sucursal:='';
                    sNro_caja:='';
                    sPeriodo_caja:='';
                    for i := 0 to slSentencias_comunicacion.Count-1 do
                    begin
                      if LeftStr(slSentencias_comunicacion[i],34)= 'INSERT INTO comprobantes_de_ventas' then
                      begin
                        sSentencia:= slSentencias_comunicacion[i];
                        sSentencia:= AnsiReplaceStr(sSentencia,'_latin1','');

                        iPosicion_caracter_inicial:=PosEx('VALUES (', sSentencia);
                        if iPosicion_caracter_inicial> -1 then
                        begin
                          iPosicion_caracter_inicial := iPosicion_caracter_inicial + 8;
                          sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, length(sSentencia)-iPosicion_caracter_inicial+1);

                          // Empresa
                          iPosicion_caracter_final:=PosEx(',',sSentencia);
                          if iPosicion_caracter_final> -1 then
                          begin
                            sId_empresa:= Trim(copy(sSentencia,0, (iPosicion_caracter_final-1)));
                          end;

                          // Siguiente y salteo un parametro
                          for j := 1 to 2 do
                          begin
                            iPosicion_caracter_inicial:=PosEx(','+Chr(13), sSentencia);
                            if iPosicion_caracter_inicial >= -1 then
                            begin
                              iPosicion_caracter_inicial:= iPosicion_caracter_inicial + 1;
                              sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, length(sSentencia)-iPosicion_caracter_inicial+1);
                            end;
                          end;

                          // Sucursal
                          iPosicion_caracter_final:=PosEx(',',sSentencia);
                          if iPosicion_caracter_final> -1 then
                          begin
                            sId_Sucursal:= Trim(copy(sSentencia,0, (iPosicion_caracter_final-1)));
                          end;

                          // Siguiente
                          for j := 1 to 1 do
                          begin
                            iPosicion_caracter_inicial:=PosEx(','+Chr(13), sSentencia);
                            if iPosicion_caracter_inicial >= -1 then
                            begin
                              iPosicion_caracter_inicial:= iPosicion_caracter_inicial + 1;
                              sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, length(sSentencia)-iPosicion_caracter_inicial+1);
                            end;
                          end;

                          // Nro_caja
                          iPosicion_caracter_final:=PosEx(',',sSentencia);
                          if iPosicion_caracter_final> -1 then
                          begin
                            sNro_caja:= Trim(copy(sSentencia,0, (iPosicion_caracter_final-1)));
                          end;

                          // Siguiente y salteo un parametro
                          for j := 1 to 2 do
                          begin
                            iPosicion_caracter_inicial:=PosEx(','+Chr(13), sSentencia);
                            if iPosicion_caracter_inicial >= -1 then
                            begin
                              iPosicion_caracter_inicial:= iPosicion_caracter_inicial + 1;
                              sSentencia:= copy(sSentencia, iPosicion_caracter_inicial, length(sSentencia)-iPosicion_caracter_inicial+1);
                            end;
                          end;

                          // Periodo_caja
                          iPosicion_caracter_final:=PosEx(',',sSentencia);
                          if iPosicion_caracter_final> -1 then
                          begin
                            sPeriodo_caja:= Trim(copy(sSentencia,0, (iPosicion_caracter_final-1)));
                          end;

                        end;

                        Break;
                      end;
                    end;



                    if ((Length(sId_empresa)>0) and (Length(sId_Sucursal)>0) and (Length(sNro_caja)>0) and (Length(sPeriodo_caja)>0)) then
                    begin
                      try

                        qrySentencias_comunicacion.SQL.Clear;

                        slSentenciasManuales:= _Confecciona_Sentencias_Faltantes('registro de cajas', sId_empresa, sId_Sucursal, sNro_caja, sPeriodo_caja, pSucursal);
                        if Assigned(slSentenciasManuales) then
                        begin
                          if slSentenciasManuales.Count>0 then
                          begin
                            qrySentencias_comunicacion.Connection.BeginTrans;
                            for i:=0 to slSentenciasManuales.Count-1 do
                            begin
                              qrySentencias_comunicacion.SQL.Clear;
                              qrySentencias_comunicacion.SQL.Add(slSentenciasManuales[i]);
                              qrySentencias_comunicacion.ExecSQL;
                            end;

                            qrySentencias_comunicacion.Connection.CommitTrans;

                            sCadena_error:= chr(13) + 'Registro de caja Registrado y corregido: Empresa '+sId_empresa+' - Sucursal '+ sId_Sucursal+' - Nro. Caja '+sNro_caja+' Periodo '+sPeriodo_caja;
                            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);

                            sCadena_error:= '';
                            if Assigned (slSentencias_comunicacion) then
                              slSentencias_comunicacion.Clear;

                            FreeAndNil(slSentenciasManuales);

                            pSucursal.GraboError           := False;
                            pSucursal.GraboError2Instancia := False;

                            sCadena_error:= '';
                          end;
                        end;
                      except
                        on E: Exception do
                        begin

                          if qrySentencias_comunicacion.Connection.InTransaction then
                            qrySentencias_comunicacion.Connection.RollbackTrans;

                          sCadena_error:= 'No se pudo registrar y corregir Registro de caja Registrado y corregido: Empresa '+sId_empresa+' - Sucursal '+ sId_Sucursal+' - Nro. Caja '+sNro_caja+' Periodo '+sPeriodo_caja + Chr(13) + '('+IntToStr(E.HelpContext)+')'+(E.Message)+ Chr(13)+'SQL: ';
                          if Assigned(slSentenciasManuales) then
                          begin
                            for i:=0 to slSentenciasManuales.Count-1 do
                            begin
                              sCadena_error:= sCadena_error + slSentenciasManuales[i] +Chr(13);
                            end;
                            FreeAndNil(slSentenciasManuales);
                          end;
                          if pSucursal.GraboError2Instancia = False then
                          begin
                            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
                            pSucursal.GraboError2Instancia:= true;
                            _CargarItemsMenu;
                          end;
                        end;
                      end;
                      qrySentencias_comunicacion.SQL.Clear;
                    end
                    else
                    begin
                      if pSucursal.GraboError2Instancia = False then
                      begin
                        sCadena_error:= 'No se pudo registrar y corregir Registro de caja Registrado y corregido: Empresa '+sId_empresa+' - Sucursal '+ sId_Sucursal+' - Nro. Caja '+sNro_caja+' Periodo '+sPeriodo_caja;
                        _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
                        pSucursal.GraboError2Instancia:= true;
                        _CargarItemsMenu;
                      end;
                    end;
                  end;
                end

                /////////////////////////////////////////////////
                // Saltea
                //
                /////////////////////////////////////////////////
                else if ((PosEx('3---Sucursal:', sCadena_error)> 0)
                  and (PosEx('You can'+Chr(39)+'t specify target table '+QuotedStr('articulos')+' for update in FROM clause', sCadena_error)> 0)
                  and (PosEx('UPDATE articulos SET articulos.excluir_actu_pcios_sud=0 WHERE id_articulo IN(SELECT articulos.id_articulo FROM articulos  INNER JOIN articulos_dto_por_multiplo ON articulos.`id_articulo`=articulos_dto_por_multiplo.`id_articulo` WHERE', sCadena_error)> 0)) then
                begin

                  // Nueva forma de sacar el End_log_pos porque en mariadb windows viene en 0 cuando no es el final de la transaccion
                  if FieldByName('End_log_pos').AsInteger>0 then
                    iEnd_log_pos_real:= FieldByName('End_log_pos').AsInteger
                  else
                  begin
                    if RecNo+1 <= RecordCount then
                    begin
                      Next;
                      iEnd_log_pos_real:= FieldByName('pos').AsInteger;
                      Prior;
                    end
                    else
                      iEnd_log_pos_real:= FieldByName('pos').AsInteger;
                  end;
                  //

                  // saltear sentencia
                  qryUpdate_FileandPos_BDDestino.Connection.BeginTrans;

                  qryUpdate_FileandPos_BDDestino.Parameters.ParamByName('pId_archivo_log_procesado').Value  := FieldByName('Log_name').AsString;
                  qryUpdate_FileandPos_BDDestino.Parameters.ParamByName('pPos_log_procesado').Value         := iEnd_log_pos_real;
                  qryUpdate_FileandPos_BDDestino.Parameters.ParamByName('pId_sucursal').Value               := pSucursal.Id_sucursal;
                  qryUpdate_FileandPos_BDDestino.ExecSQL;

                  qryUpdate_FileandPos_BDDestino.Connection.CommitTrans;

                  //Actualizo el objeto sucursal que tengo en memoria con los ultimos Log_name y pos que logre actualizar la BD destino
                  //Blanqueo la variable de grabacion del error
                  pSucursal.Id_archivo_log_procesado  := FieldByName('Log_name').AsString;
                  pSucursal.Pos_log_procesado         := iEnd_log_pos_real;
                  pSucursal.GraboError                := False;
                  pSucursal.GraboError2Instancia      := False;
                  sCadena_error:= '';
                  if Assigned (slSentencias_comunicacion) then
                    slSentencias_comunicacion.Clear;
                  qrySentencias_comunicacion.SQL.Clear;
                  iPosIni:= -1;

                end
                /////////////////////////////////////////////////
                // Saltea
                //
                /////////////////////////////////////////////////
                else if ((PosEx('3---Sucursal:', sCadena_error)> 0)
                  and (PosEx('Cannot delete or update a parent row: a foreign key constraint fails', sCadena_error)> 0)
                  and (PosEx('`pedidos`, CONSTRAINT `pedidos_ibfk_11` FOREIGN KEY (`id_empresa`, `id_comprobante`) REFERENCES `comprobantes_de_ventas`', sCadena_error)> 0)) then
                begin

                  // Nueva forma de sacar el End_log_pos porque en mariadb windows viene en 0 cuando no es el final de la transaccion
                  if FieldByName('End_log_pos').AsInteger>0 then
                    iEnd_log_pos_real:= FieldByName('End_log_pos').AsInteger
                  else
                  begin
                    if RecNo+1 <= RecordCount then
                    begin
                      Next;
                      iEnd_log_pos_real:= FieldByName('pos').AsInteger;
                      Prior;
                    end
                    else
                      iEnd_log_pos_real:= FieldByName('pos').AsInteger;
                  end;
                  //

                  // saltear sentencia
                  qryUpdate_FileandPos_BDDestino.Connection.BeginTrans;

                  qryUpdate_FileandPos_BDDestino.Parameters.ParamByName('pId_archivo_log_procesado').Value  := FieldByName('Log_name').AsString;
                  qryUpdate_FileandPos_BDDestino.Parameters.ParamByName('pPos_log_procesado').Value         := iEnd_log_pos_real;
                  qryUpdate_FileandPos_BDDestino.Parameters.ParamByName('pId_sucursal').Value               := pSucursal.Id_sucursal;
                  qryUpdate_FileandPos_BDDestino.ExecSQL;

                  qryUpdate_FileandPos_BDDestino.Connection.CommitTrans;

                  //Actualizo el objeto sucursal que tengo en memoria con los ultimos Log_name y pos que logre actualizar la BD destino
                  //Blanqueo la variable de grabacion del error
                  pSucursal.Id_archivo_log_procesado  := FieldByName('Log_name').AsString;
                  pSucursal.Pos_log_procesado         := iEnd_log_pos_real;
                  pSucursal.GraboError                := False;
                  pSucursal.GraboError2Instancia      := False;
                  sCadena_error:= '';
                  if Assigned (slSentencias_comunicacion) then
                    slSentencias_comunicacion.Clear;
                  qrySentencias_comunicacion.SQL.Clear;
                  iPosIni:= -1;

                end
                else
                begin
                  if Assigned (slSentencias_comunicacion) then
                    slSentencias_comunicacion.Clear;
                  Exit;
                end;
                //////////////////////////////////////////////////////////
                //Exit;
              end;
            end;

          end
          else
          begin

            bEsAnnotate_rows:= False;
            if ((((PosEx('UPDATE `', FieldByName('info').AsString)>0)      and (PosEx('`articulos` SET', FieldByName('info').AsString)>0))
                or ((PosEx('INSERT INTO `', FieldByName('info').AsString)>0) and (PosEx('`articulos`(', FieldByName('info').AsString)>0))
                or ((PosEx('DELETE FROM `', FieldByName('info').AsString)>0) and (PosEx('`articulos`', FieldByName('info').AsString)>0))) and ( (LowerCase(FieldByName('Event_type').AsString) = 'annotate_rows')) and (bBD_erronea= False)) then // Esto es por la actualizacion Masiva de precios
            begin
              bEsAnnotate_rows:= True;
            end;

            // Si es COMMIT y no es la BD, actualizo el log
            if (((LeftStr(FieldByName('info').AsString, 6 ) = 'COMMIT') and ( (LowerCase(FieldByName('Event_type').AsString) = 'xid') or (LowerCase(FieldByName('Event_type').AsString) = 'query') or (LowerCase(FieldByName('Event_type').AsString) = 'annotate_rows')) and (bBD_erronea= True)) or bEsAnnotate_rows) then
            begin

              // Nueva forma de sacar el End_log_pos porque en mariadb windows viene en 0 cuando no es el final de la transaccion
              if FieldByName('End_log_pos').AsInteger>0 then
                iEnd_log_pos_real:= FieldByName('End_log_pos').AsInteger
              else
              begin
                if RecNo+1 <= RecordCount then
                begin
                  Next;
                  iEnd_log_pos_real:= FieldByName('pos').AsInteger;
                  Prior;
                end
                else
                  iEnd_log_pos_real:= FieldByName('pos').AsInteger;
              end;
              //

              pSucursal.Id_archivo_log_procesado  := FieldByName('Log_name').AsString;
              pSucursal.Pos_log_procesado         := iEnd_log_pos_real;
              iPosIni                             := iEnd_log_pos_real;
            end;

            //if LowerCase(FieldByName('Event_type').AsString) =  'query' then
            if ((LowerCase(FieldByName('Event_type').AsString)='query') or (LowerCase(FieldByName('Event_type').AsString) = 'annotate_rows') or (LowerCase(FieldByName('Event_type').AsString)='gtid')) then
            begin
              sCadena:= FieldByName('info').AsString;

              iPosicion_nombreBD:= 5;      //use `zafiro`; UPDATE cupones_promocion
              if LowerCase(LeftStr(sCadena, 11)) = 'insert into' then //INSERT INTO `zafiro`.`marcas`(`id_marca`,`
                iPosicion_nombreBD:=13
              else if LowerCase(LeftStr(sCadena, 11)) = 'delete from' then  // DELETE FROM `zafiro`.`marcas` WHERE `id_marca`='0'
                iPosicion_nombreBD:=13
              else if LowerCase(LeftStr(sCadena, 6)) = 'update' then  // UPDATE `zafiro`.`articulos` SET
                iPosicion_nombreBD:=8;

              //Saco el nombre de la BD de la sentencia para comparar si se trata de la misma BD con la que estoy trabajando es decir la que se encuentra en la tabla parámetros
              sNombreBD:='';
              while iPosicion_nombreBD<=(Length(sCadena)) do //mientras que no se llegue a la posición final de la cadena
              begin
                //IMPORTANTE: esta sentencia esta muy sujeta a como guarda el MySql los query, por lo que puede variar en otra configuración
                if (sCadena[iPosicion_nombreBD]='`') then
                begin
                  //Si ya encontre el nombre de la BD salgo del bucle
                  if Length(sNombreBD) > 0 then
                    Break;
                end
                else
                begin
                  sNombreBD:= sNombreBD + sCadena[iPosicion_nombreBD];
                end;
                iPosicion_nombreBD := iPosicion_nombreBD + 1;
              end; //end while pPosicion_caracter_final<=(Length(pCadena)) do //mientras que no se llegue a la posición final de la cadena


              if ((oParametro.Nombre_BD = sNombreBD) or (PosEx('begin', LowerCase(sCadena))> 0)) then
              //if ((oParametro.Nombre_BD = sNombreBD) or (LeftStr(sCadena, 5 ) = 'BEGIN')) then
              begin
                bBD_erronea:= False;

                //*********Identifico si se trata de un insert o delete o update para guardar la cadena en el sql. Tambien un Begin para indicar iPosIni**********//
                //PosEx me devuelve la posición del primer caracter de una subcadena especificada
                if ((PosEx('insert', LowerCase(sCadena))> 0) and (PosEx('into', LowerCase(sCadena))> 0 ))then
                begin
                  sCadena:= RightStr(sCadena,(Length(sCadena) - (PosEx('insert', LowerCase(sCadena))-1)) );
                  iPosicion_caracter_inicial:=PosEx('into', LowerCase(sCadena));
                  if iPosicion_caracter_inicial> 0  then
                  begin
                    iPosicion_caracter_final:= iPosicion_caracter_inicial + 4;
                    _Agregar_CadenaSQL(pSucursal,sCadena, iPosicion_caracter_final, pTabla_Exportar);
                    // Verifico si es un comprobante temporal
                    if LowerCase(_Extraer_NombreTabla(sCadena, iPosicion_caracter_final))='comprobantes_de_ventas' then
                    begin
                      if ((PosEx('_latin1'+Chr(39)+'TE', sCadena)>0)) then  // id_comprobante  //  or (PosEx('_latin1'+Chr(39)+'TR', sCadena)>0)
                        if ((PosEx('_latin1'+Chr(39)+'TE'+Chr(39)+',', sCadena)>0) or (PosEx('_latin1'+Chr(39)+'TG'+Chr(39)+',', sCadena) > 0)) then  // estado
                          if ((PosEx('_latin1'+Chr(39)+'FAC'+Chr(39)+',', sCadena) > 0) or (PosEx('_latin1'+Chr(39)+'ORD'+Chr(39)+',', sCadena) > 0) or (PosEx('_latin1'+Chr(39)+'NC'+Chr(39)+',', sCadena) > 0)) then  // Tipo Comprobante
                            // La cadena *R@D se registra en 'origen_preventa' - //  Cadena que indica que es preventa redireccionada
                            if (PosEx('*R@D', sCadena) = 0) then  // Si es 0, quiere decir que es una preventa normal no redireccionada a otra sucursal
                              bComprobanteTemporal:= True;
                      // Para saltear las ordenes
                      //if (PosEx('_latin1'+Chr(39)+'CI'+Chr(39), sCadena) > 0) then
                      //  if (PosEx('_latin1'+Chr(39)+'ORD'+Chr(39)+',', sCadena) > 0) then
                      //    if (PosEx('_latin1'+Chr(39)+'EM'+Chr(39)+',', sCadena) > 0) then
                      //      bComprobanteTemporal:= True;
                      //

                      // Si es UPCN y es temporal, no excluyo los temporales guardados
                      if bComprobanteTemporal then
                      begin
                        if oParametro.Importacion_UPCN  then
                        begin
                          if ((PosEx('_latin1'+Chr(39)+'TG'+Chr(39)+',', sCadena) > 0)) then  // estado
                          begin
                            // Si es UPCN y es temporal, no excluyo los temporales guardados
                                                //bComprobanteTemporal:= False;
                            //
                          end;
                        end;
                      end;
                    end;
                    //
                  end;//end if iPosicion_caracter_inicial> 0  then
                end //end if  PosEx('insert', sCadena)> 0  then
                else if ((PosEx('delete', LowerCase(sCadena))>0) and (PosEx('from', LowerCase(sCadena))>0)) then
                begin
                  sCadena:= RightStr(sCadena,(Length(sCadena) - (PosEx('delete', LowerCase(sCadena))-1)) );
                  iPosicion_caracter_inicial:=PosEx('from', LowerCase(sCadena));
                  if iPosicion_caracter_inicial> 0  then
                  begin
                    iPosicion_caracter_final:= iPosicion_caracter_inicial + 4;
                    _Agregar_CadenaSQL(pSucursal, sCadena, iPosicion_caracter_final, pTabla_Exportar);
                  end;//end if iPosicion_caracter_inicial> 0  then

                  if LowerCase(_Extraer_NombreTabla(sCadena, iPosicion_caracter_final))='comprobantes_de_ventas' then
                  begin
                    if ((PosEx('_latin1'+Chr(39)+'TE', sCadena)>0) or (PosEx('_latin1'+Chr(39)+'TR', sCadena)>0)) then  // id_comprobante  //
                    begin
                      if (PosEx('AND '+Chr(39)+'TG'+Chr(39)+'='+Chr(39)+'TG', sCadena) > 0) then  // AND 'TG'='TG'
                        bComprobanteTemporal:= True;

                      // Es Preventa Redireccionada - Se debe trasnmitir
                      if (PosEx('AND '+Chr(39)+'TG-Redir'+Chr(39)+'='+Chr(39)+'TG-Redir', sCadena) > 0) then  // AND 'TG-Redir'='TG-Redir'
                        bComprobanteTemporal:= False;
                      //
                    end;

                  end;

                  // Si es UPCN y es temporal, no excluyo los temporales guardados
                  if bComprobanteTemporal then
                  begin
                    if oParametro.Importacion_UPCN  then
                    begin
                      if ((PosEx('_latin1'+Chr(39)+'TG'+Chr(39)+',', sCadena) > 0)) then  // estado
                      begin
                        // Si es UPCN y es temporal, no excluyo los temporales guardados
                                               //bComprobanteTemporal:= False;
                        //
                      end;
                    end;
                  end;


                end //end if  PosEx('delete', sCadena)> 0  then
                else if ((PosEx('update', LowerCase(sCadena))>0) and (PosEx('set', LowerCase(sCadena))>0)) then
                begin
                  iPosicion_caracter_inicial:=PosEx('update', LowerCase(sCadena));
                  if  iPosicion_caracter_inicial> 0  then
                  begin
                    sCadena:= RightStr(sCadena,(Length(sCadena) - (iPosicion_caracter_inicial-1)) );
                    if PosEx('set', LowerCase(sCadena))> 0  then
                    begin

                      bExcluyeCadena:= False;
                      // Error en Farmanor
                      if (PosEx('UPDATE articulos SET articulos.excluir_actu_pcios_sud=0 WHERE id_articulo IN(SELECT articulos.id_articulo FROM articulos', sCadena)> 0) then
                        bExcluyeCadena:= True;

                      if (PosEx('UPDATE pedidos LEFT JOIN comprobantes_de_ventas ON (pedidos.id_empresa=comprobantes_de_ventas.id_empresa AND pedidos.id_pedido=comprobantes_de_ventas.id_pedido) SET pedidos.id_comprobante=comprobantes_de_ventas.id_comprobante', sCadena)> 0) then
                        bExcluyeCadena:= True;

                      // Para que no transmita los update del BI
                      if (PosEx('UPDATE comprobantes_de_ventas SET comprobantes_de_ventas.z_bi = 1 WHERE comprobantes_de_ventas.id_empresa =', sCadena)> 0) then
                        bExcluyeCadena:= True;
                      if (PosEx('UPDATE articulos SET articulos.z_bi =', sCadena)> 0) then
                        bExcluyeCadena:= True;
                      if (PosEx('UPDATE comprobantes_mov_stock SET comprobantes_mov_stock.z_bi = 1 WHERE comprobantes_mov_stock.id_empresa =', sCadena)> 0) then
                        bExcluyeCadena:= True;
                      if (PosEx('UPDATE proveedores SET proveedores.z_bi =', sCadena)> 0) then
                        bExcluyeCadena:= True;
                      //
                      // Para que no transmita los update talonarios nro_proximo
                      if (PosEx('UPDATE talonarios SET nro_proximo =', sCadena)> 0) then
                        bExcluyeCadena:= True;
                      //

                      // Para que no transmita ataque en La Esquina
                      if ((PosEx('UPDATE preasientos_contables LEFT JOIN comprobantes_de_ventas ON (preasientos_contables.id_empresa=comprobantes_de_ventas.id_empresa AND preasientos_contables.id_comprobante=comprobantes_de_ventas.id_comprobante)', sCadena)> 0)
                          and (PosEx('SET preasientos_contables.nro_asiento= 5386', sCadena)> 0)
                          and (PosEx('WHERE preasientos_contables.nro_asiento IS NULL', sCadena)> 0)
                          and (PosEx('AND comprobantes_de_ventas.id_empresa   = 6', sCadena)> 0)
                          and (PosEx('AND comprobantes_de_ventas.id_sucursal  = 15', sCadena)> 0)
                          and (PosEx('AND comprobantes_de_ventas.nro_caja     = 42', sCadena)> 0)
                          and (PosEx('AND comprobantes_de_ventas.periodo_caja = 6', sCadena)> 0) ) then
                        bExcluyeCadena:= True;
                      //

                      // Para que no transmita ataque en La Esquina
                      if ((PosEx('UPDATE preasientos_contables LEFT JOIN comprobantes_de_ventas ON (preasientos_contables.id_empresa=comprobantes_de_ventas.id_empresa AND preasientos_contables.id_comprobante=comprobantes_de_ventas.id_comprobante)', sCadena)> 0)
                          and (PosEx('SET preasientos_contables.nro_asiento= 6109', sCadena)> 0)
                          and (PosEx('WHERE preasientos_contables.nro_asiento IS NULL', sCadena)> 0)
                          and (PosEx('AND comprobantes_de_ventas.id_empresa   = 3', sCadena)> 0)
                          and (PosEx('AND comprobantes_de_ventas.id_sucursal  = 5', sCadena)> 0)
                          and (PosEx('AND comprobantes_de_ventas.nro_caja     = 26', sCadena)> 0)
                          and (PosEx('AND comprobantes_de_ventas.periodo_caja = 56', sCadena)> 0) ) then
                        bExcluyeCadena:= True;
                      //

                      // Para que no transmita ataque en La Esquina
                      if ((PosEx('UPDATE preasientos_contables LEFT JOIN comprobantes_de_ventas ON (preasientos_contables.id_empresa=comprobantes_de_ventas.id_empresa AND preasientos_contables.id_comprobante=comprobantes_de_ventas.id_comprobante)', sCadena)> 0)
                          and (PosEx('SET preasientos_contables.nro_asiento= 78514', sCadena)> 0)
                          and (PosEx('WHERE preasientos_contables.nro_asiento IS NULL', sCadena)> 0)
                          and (PosEx('AND comprobantes_de_ventas.id_empresa   = 4', sCadena)> 0)
                          and (PosEx('AND comprobantes_de_ventas.id_sucursal  = 6', sCadena)> 0)
                          and (PosEx('AND comprobantes_de_ventas.nro_caja     = 72', sCadena)> 0)
                          and (PosEx('AND comprobantes_de_ventas.periodo_caja = 78', sCadena)> 0) ) then
                        bExcluyeCadena:= True;
                      //


                      if bExcluyeCadena=False then
                      begin

                        bEsSucursales_comunicacionFecha_Hora:= False;
                        // Transforma la sentencia para registrar la fecha y la hora de transmision local y en destino
                        if PosEx('UPDATE sucursales_comunicacion SET sucursales_comunicacion.fecha_hora', sCadena) > 0 then
                        begin

                          bEsSucursales_comunicacionFecha_Hora:= True;
                          sSentencia:= sCadena;
                          sSentencia := AnsiReplaceStr(sSentencia,'sucursales_comunicacion','sucursales_log_tiempo_comunicacion');
                          sSentencia := AnsiReplaceStr(sSentencia,'fecha_hora','envio_fecha_hora_comunicacion');
                          sSentencia := AnsiReplaceStr(sSentencia,'.id_sucursal = '+IntToStr(oParametro.Comu_Id_Sucursal) ,'.id_sucursal = '+IntToStr(pSucursal.Id_sucursal));


                          qryUpdate_sucursales_log_tiempo_comunicacion_local.SQL.Clear;
                          qryUpdate_sucursales_log_tiempo_comunicacion_local.Connection.BeginTrans;
                          qryUpdate_sucursales_log_tiempo_comunicacion_local.SQL.Add(sSentencia);
                          qryUpdate_sucursales_log_tiempo_comunicacion_local.ExecSQL;
                          qryUpdate_sucursales_log_tiempo_comunicacion_local.SQL.Clear;
                          qryUpdate_sucursales_log_tiempo_comunicacion_local.Connection.CommitTrans;

                          //_Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena+Chr(13)+sSentencia+Chr(13));

                          sSentencia := AnsiReplaceStr(sSentencia,'envio_fecha_hora_comunicacion','recepcion_fecha_hora_comunicacion');
                          iPosicion_caracter_inicial:= PosEx('WHERE sucursales_log_tiempo_comunicacion.id_sucursal =', sSentencia);
                          sSentencia := LeftStr(sSentencia,iPosicion_caracter_inicial-1)+ 'WHERE sucursales_log_tiempo_comunicacion.id_sucursal = '+IntToStr(oParametro.Comu_Id_Sucursal);
                          sCadena := sSentencia;

                        end;

                        // Para omitir la sentencia que actualiza la sucursales_log_tiempo_comunicacion
                        //if LeftStr(sCadena,112)='UPDATE sucursales_log_tiempo_comunicacion SET sucursales_log_tiempo_comunicacion.envio_fecha_hora_comunicacion =' then
                        if ((PosEx('sucursales_log_tiempo_comunicacion', sCadena)>0) and (PosEx('envio_fecha_hora_comunicacion', sCadena)>0)) then
                        begin
                          // Para que no se transmita este UPDATE
                          bExcluyeCadena := True;
                          //
                        end;
                        //

                        if bEsSucursales_comunicacionFecha_Hora= False then
                        begin
                          // Para omitir la sentencia que actualiza la sucursales_log_tiempo_comunicacion
                          //if LeftStr(sCadena,112)='UPDATE sucursales_log_tiempo_comunicacion SET sucursales_log_tiempo_comunicacion.envio_fecha_hora_comunicacion =' then
                          if ((PosEx('sucursales_log_tiempo_comunicacion', sCadena)>0) and (PosEx('recepcion_fecha_hora_comunicacion', sCadena)>0)) then
                          begin
                            // Para que no se transmita este UPDATE para las otras sucursales
                            bExcluyeCadena := True;
                            //
                          end;
                          //
                        end;
                      end;

                      if bExcluyeCadena=False then
                      begin
                        _Agregar_CadenaSQL(pSucursal, sCadena, 7, pTabla_Exportar);
                      end;

                      if LowerCase(_Extraer_NombreTabla(sCadena, iPosicion_caracter_final))='comprobantes_de_ventas' then
                      begin
                        if ((PosEx('_latin1'+Chr(39)+'TE', sCadena)>0) or (PosEx('_latin1'+Chr(39)+'TR', sCadena)>0)) then  // id_comprobante  //
                          if ((PosEx('SET editado = 1', sCadena) > 0) or (PosEx('SET editado = 0', sCadena) > 0)) then
                            // bComprobanteTemporal:= True;
                      end;

                      // Si es UPCN y es temporal, no excluyo los temporales guardados
                      if bComprobanteTemporal then
                      begin
                        if oParametro.Importacion_UPCN  then
                        begin
                          if ((PosEx('_latin1'+Chr(39)+'TG'+Chr(39)+',', sCadena) > 0)) then  // estado
                          begin
                            // Si es UPCN y es temporal, no excluyo los temporales guardados
                                                   //bComprobanteTemporal:= False;
                            //
                          end;
                        end;
                      end;

                    end;//end if if PosEx('set', sCadena)> 0  then
                  end; //end if iPosicion_caracter_inicial> 0  then
                end //if PosEx('update', LowerCase(sCadena))> 0  then
                else if PosEx('begin', LowerCase(sCadena))> 0  then
                begin
                  bComprobanteTemporal:= False;

                  if slSentencias_comunicacion.Count = 0 then // no hay ninguna sentencia previa
                  begin
                    if iPosIni<0 then
                      iPosIni:= FieldByName('pos').AsInteger;
                  end;
                end; //if PosEx('update', LowerCase(sCadena))> 0  then

                // crisis
                //      iPosIni:= FieldByName('pos').AsInteger;   // Agregado prisis
                // crisis fin


                //*********Identifico si se trata de un insert o delete o update para guardar la cadena en el sql. Tambien un Begin para indicar iPosIni**********//

              end
              else
              begin
                bBD_erronea:= True;

                if PosEx('begin', LowerCase(sCadena))> 0 then
                begin
                  if oParametro.Nombre_BD <> sNombreBD then
                  begin
                    // Si no es la BD, adelanto el iPosIni
                    if slSentencias_comunicacion.Count = 0 then // no hay ninguna sentencia previa
                    begin
                      if iPosIni<0 then
                        iPosIni:= FieldByName('pos').AsInteger;
                    end;
                    //
                  end;
                end;
              end;//end if oParametro.Nombre_BD = sNombreBD then

            end
            else
            begin
              if ((LowerCase(FieldByName('Event_type').AsString) =  'stop') or (LowerCase(FieldByName('Event_type').AsString) =  'rotate')) then
              begin
                bFin_archivo:= True;
              end;

            end; //end if   LowerCase(FieldByName('Event_type').AsString) =  'query' then

          end; // if  LeftStr(FieldByName('info').AsString, 6 ) = 'COMMIT' then
          Next;

//          if bBD_erronea then
//            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, 'Next' + '  '+ FieldByName('pos').AsString+ ' Erronea')
//          else
//            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, 'Next' + '  '+ FieldByName('pos').AsString + ' Ok');


        end; // end while Not Eof do

      end;

      if bFin_archivo then
      begin
        ///////////////////////////////////////////////////////////////////////////////
        //  Esto es lo que generaba saltos en la comunicacion !!!!                   //
        //  Inicio                                                                   //
        //  Ahora se hace solo si llego al fin del archovo log -> bFin_archivo=True  //
        ///////////////////////////////////////////////////////////////////////////////
        try

//          sCadena_error:= 'Id_archivo:'+  pSucursal.Id_archivo_log_procesado + ' Pos: '+ IntToStr(iEnd_log_pos) + ' Sucursal:'+ IntToStr(pSucursal.Id_sucursal);
//          _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);

          qryUpdate_FileandPos_BDDestino.Connection.BeginTrans;

          qryUpdate_FileandPos_BDDestino.Parameters.ParamByName('pId_archivo_log_procesado').Value  := pSucursal.Id_archivo_log_procesado;
          qryUpdate_FileandPos_BDDestino.Parameters.ParamByName('pPos_log_procesado').Value         := iEnd_log_pos;
          qryUpdate_FileandPos_BDDestino.Parameters.ParamByName('pId_sucursal').Value               := pSucursal.Id_sucursal;
          qryUpdate_FileandPos_BDDestino.ExecSQL;

          qryUpdate_FileandPos_BDDestino.Connection.CommitTrans;
          //qrySentencias_comunicacion.SQL.Clear; //Limpio el sql para que siga cargando

          //Actualizo el objeto sucursal que tengo en memoria con los ultimos Log_name y pos que logre actualizar la BD destino
          pSucursal.Pos_log_procesado         := iEnd_log_pos;
        except
          on E: Exception do
          begin
            if qryUpdate_FileandPos_BDDestino.Connection.InTransaction then
              qryUpdate_FileandPos_BDDestino.Connection.RollbackTrans;
            sCadena_error:= DateTimeToStr(Now)+'  4---Sucursal:'+inttostr(pSucursal.Id_sucursal)+' - '+pSucursal.Des_sucursal+' -  Error: '+ (E.Message)+
                                                Chr(13)+'SQL: '+ qryUpdate_FileandPos_BDDestino.SQL.Text;

            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);

            Exit;
          end;
        end;//end try
        ///////////////////////////////////////////////////////////////////////////////
        //  Esto es lo que generaba saltos en la comunicacion !!!!                   //
        //  Fin                                                                      //
        //  Ahora se hace solo si llego al fin del archovo log -> bFin_archivo=True  //
        ///////////////////////////////////////////////////////////////////////////////
      end;

    end; //end if RecordCount > 0 then
  end;//end with qrySelect_Binlog_events do

  if Assigned (slSentencias_comunicacion) then
    slSentencias_comunicacion.Free;

  // Registro de fecha hora para log de transmisión
//  if ((Self.Tag=1) or (Self.Tag=2) or (Self.Tag=3)) then   // Solo en el tag = 1 se ejecuta. Osea en un solo servicio de tag 1
//  begin
    sCadena_error:= _Modificar_Comu_Fecha_Hora(oParametro.Comu_Id_Sucursal);
    if Length(sCadena_error)>0 then
    begin
      sCadena_error:= DateTimeToStr(Now)+'  14---Sucursal:'+IntToStr(pSucursal.Id_sucursal)+' - '+pSucursal.Des_sucursal+' - '+ sCadena_error;
      _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);
    end;

//  end;
  //

  //if (((pSucursal.Id_archivo_log_procesado = sUltimo_archivo_log_BDorigen) and (pSucursal.Pos_log_procesado = iEnd_log_pos)) or (pSucursal.Id_archivo_log_procesado > sUltimo_archivo_log_BDorigen)) then
  if (((pSucursal.Id_archivo_log_procesado = sUltimo_archivo_log_BDorigen) and (pSucursal.Pos_log_procesado >= iUltimo_Pos_log_BDorigen))
    or (pSucursal.Id_archivo_log_procesado > sUltimo_archivo_log_BDorigen)) then
  begin

    // Llego al final del log.
    _Transmision_de_Sentencias_Por_Comprobante(pSucursal);
    //

    Result:= False;
    Exit;
  end
  else
  begin
    if _Existe_Mas_Sentencias_del_mismo_log(pSucursal.Id_archivo_log_procesado, pSucursal.Pos_log_procesado) then
    begin
      // Hay mas el el mismo archivo log
    end
    else
    begin
      //Si termino el bucle anterior, no hay mas para el mismo log y no igualo a la bd origen significa que hay
      //que seguir abriendo los archivos siguientes x lo que incremento el Id_archivo_log_procesado e
      //inicializo Pos_log_procesado para que se lea el archivo desde el principio

      if bFin_archivo = False then
        bFin_archivo:= _Existe_log_Posterior(pSucursal.Id_archivo_log_procesado);

      if bFin_archivo = True then
      begin
        pSucursal.Pos_log_procesado:= 0;
        iNro_Extension:= StrToInt(RightStr(pSucursal.Id_archivo_log_procesado, oParametro.Comu_largo_extension_file_binlog)) + 1;
        pSucursal.Id_archivo_log_procesado:= oParametro.Comu_nombre_file_binlog + '.'+ RightStr('00000000000000000'+IntToStr(iNro_Extension), oParametro.Comu_largo_extension_file_binlog);

        try
          qryUpdate_FileandPos_BDDestino.Connection.BeginTrans;

          qryUpdate_FileandPos_BDDestino.Parameters.ParamByName('pId_archivo_log_procesado').Value  := pSucursal.Id_archivo_log_procesado;
          qryUpdate_FileandPos_BDDestino.Parameters.ParamByName('pPos_log_procesado').Value         := pSucursal.Pos_log_procesado;
          qryUpdate_FileandPos_BDDestino.Parameters.ParamByName('pId_sucursal').Value               := pSucursal.Id_sucursal;
          qryUpdate_FileandPos_BDDestino.ExecSQL;

          qryUpdate_FileandPos_BDDestino.Connection.CommitTrans;
        except
          on E: Exception do
          begin
            if qryUpdate_FileandPos_BDDestino.Connection.InTransaction then
              qryUpdate_FileandPos_BDDestino.Connection.RollbackTrans;
            sCadena_error:= DateTimeToStr(Now)+'  5---Sucursal:'+IntToStr(pSucursal.Id_sucursal)+' - '+pSucursal.Des_sucursal+' -  Error: '+ (E.Message)+
                                                Chr(13)+'SQL: '+qryUpdate_FileandPos_BDDestino.SQL.Text;

            _Grabar_LogErrores_Comunicacion(pSucursal.Id_ProcesoActualizacion, sCadena_error);

            Exit;
          end;
        end;
      end;
    end;




    //_ActualizarBD(pSucursal);
    if pSucursal.ErrorVerificacionSentencias=0 then
      Result:= True;
  end;
end;

procedure TVSComunicacionParaReplicacion01._Actualizar_Min_Max_Stock(pId_Empresa: Integer);
var
  dsStock_min_sugerido  : TDataSet;

  dsSucursales: TDataSet;
  cdsArticulos : TClientDataSet;
  sId_Articulo, sFecha_desde, sFecha_hasta  : String;
  iDiasPeriodoAnalisis, iDiasDeStock:  Integer;
  dCoefStockMaximo: Double;
  dtFechaDesde, dtFechaHasta : TDate;
  sMensaje_error, sFiltro_Sucursales : String;
  dsArticulosConsulta  : TDataSet;
  iId_sucursal : Integer;
  oBarraProgreso : TProgressBar;
  oNivel_de_Rotacion: TNiveles_de_Rotacion;
begin
  inherited;
  dsArticulosConsulta := nil;
  dsArticulosConsulta := _Gestor_Articulos._Get_Articulos_Actualizacion_Automatica_Min_Max_Stock(pId_Empresa);

  if not Assigned(_Gestor_Niveles_Rotacion) then
    _Gestor_Niveles_Rotacion  := TGestor_Niveles_de_Rotacion.Create.Create;

  if Assigned(dsArticulosConsulta) then
  begin
    //Armao Dataset de Artículos para establecer stock minimo y maximo x sucursal
    cdsArticulos := TClientDataSet.Create(Nil);

    with cdsArticulos.FieldDefs.AddFieldDef do
    begin
      Name     := 'id_articulo';
      DataType := ftString;
      Size     := 255;
    end;

    with cdsArticulos.FieldDefs.AddFieldDef do
    begin
      Name     := 'des_arti_pres';
      DataType := ftString;
      Size     := 255;
    end;

    //Los campos que siguen se agregan para esta versión automatizada, no están en la pantalla de admin. stock x suc
    with cdsArticulos.FieldDefs.AddFieldDef do
    begin
      Name     := 'mym_dias_periodoanalisis_art';
      DataType := ftInteger;
    end;

    with cdsArticulos.FieldDefs.AddFieldDef do
    begin
      Name     := 'mym_dias_de_stock_art';
      DataType := ftInteger;
    end;

    with cdsArticulos.FieldDefs.AddFieldDef do
    begin
      Name     := 'mym_coefic_stockmaximo_art';
      DataType := ftFloat;
    end;


    with cdsArticulos.FieldDefs.AddFieldDef do
    begin
      Name     := 'mym_dias_periodoAnalisis_suc';
      DataType := ftInteger;
    end;

    with cdsArticulos.FieldDefs.AddFieldDef do
    begin
      Name     := 'mym_dias_de_stock_suc';
      DataType := ftInteger;
    end;

    with cdsArticulos.FieldDefs.AddFieldDef do
    begin
      Name     := 'mym_coefic_stockMaximo_suc';
      DataType := ftFloat;
    end;

    dsSucursales := nil;
    dsSucursales := _Gestor_Sucursal._Get_Todas_Sucursales(pId_Empresa);


    if Assigned(dsSucursales) then
    begin
      try
        dsSucursales.First;
        dsSucursales.Filter := 'habilita_calc_minmax_autom = 1';
        dsSucursales.Filtered := True;

        if dsSucursales.RecordCount > 0 then
        begin
          if dsSucursales.Active then
          begin
            dsSucursales.First;
            while not dsSucursales.Eof do
            begin

              with cdsArticulos.FieldDefs.AddFieldDef do
              begin
                Name     := 'id_sucursal'+IntToStr(dsSucursales.FieldByName('id_sucursal').AsInteger);
                DataType := ftInteger;
              end;

              with cdsArticulos.FieldDefs.AddFieldDef do
              begin
                Name     := 'des_sucursal'+IntToStr(dsSucursales.FieldByName('id_sucursal').AsInteger);
                DataType := ftString;
                Size     := 255;
              end;

              with cdsArticulos.FieldDefs.AddFieldDef do
              begin
                Name     := 'stock_minimo'+IntToStr(dsSucursales.FieldByName('id_sucursal').AsInteger);
                DataType := ftFloat;
              end;

              with cdsArticulos.FieldDefs.AddFieldDef do
              begin
                Name     := 'stock_maximo'+IntToStr(dsSucursales.FieldByName('id_sucursal').AsInteger);
                DataType := ftFloat;
              end;

              with cdsArticulos.FieldDefs.AddFieldDef do
              begin
                Name     := 'id_nivel_rotacion'+IntToStr(dsSucursales.FieldByName('id_sucursal').AsInteger);
                DataType := ftInteger;
              end;

              with cdsArticulos.FieldDefs.AddFieldDef do
              begin
                Name     := 'excluir_calc_aut'+IntToStr(dsSucursales.FieldByName('id_sucursal').AsInteger);
                DataType := ftString;
                Size     := 20;
              end;

              dsSucursales.Next;
            end;  // while not dsSucursales.Eof
          end;// if dsSucursales.Active
        end;// if dsSucursales.RecordCount
      finally
        // Nada
      end;
    end; // if Assigned(dsSucursales)

    with cdsArticulos.IndexDefs.AddIndexDef do
    begin
      Name := 'Principal';
      Fields := 'id_articulo';
      Options := [ixPrimary];
    end;

    cdsArticulos.CreateDataSet;
    cdsArticulos.open;

    try
      if dsArticulosConsulta.Active then
      begin
        if cdsArticulos.Active then
        begin
          iId_sucursal:=-1;
          sId_Articulo:='******************';

          dsArticulosConsulta.First;
          while not  dsArticulosConsulta.Eof do
          begin
            if dsArticulosConsulta.FieldByName('id_articulo').AsString<>sId_Articulo then
            begin
              if sId_Articulo<>'******************' then
              begin
                cdsArticulos.Post;
              end;

              cdsArticulos.Append;
              cdsArticulos.FieldByName('id_articulo').AsString    := dsArticulosConsulta.FieldByName('id_articulo').AsString;
              cdsArticulos.FieldByName('des_arti_pres').AsString  := dsArticulosConsulta.FieldByName('des_arti_pres').AsString;
              sId_Articulo:= dsArticulosConsulta.FieldByName('id_Articulo').AsString;

              cdsArticulos.FieldByName('mym_dias_periodoanalisis_art').AsInteger := dsArticulosConsulta.FieldByName('mym_dias_periodoanalisis_art').AsInteger;
              cdsArticulos.FieldByName('mym_dias_de_stock_art').AsInteger        := dsArticulosConsulta.FieldByName('mym_dias_de_stock_art').AsInteger;
              cdsArticulos.FieldByName('mym_coefic_stockmaximo_art').AsFloat     := dsArticulosConsulta.FieldByName('mym_coefic_stockmaximo_art').AsFloat;

              cdsArticulos.FieldByName('mym_dias_periodoAnalisis_suc').AsInteger := dsArticulosConsulta.FieldByName('mym_dias_periodoAnalisis_suc').AsInteger;
              cdsArticulos.FieldByName('mym_dias_de_stock_suc').AsInteger        := dsArticulosConsulta.FieldByName('mym_dias_de_stock_suc').AsInteger;
              cdsArticulos.FieldByName('mym_coefic_stockMaximo_suc').AsFloat     := dsArticulosConsulta.FieldByName('mym_coefic_stockMaximo_suc').AsFloat;

            end;

            iId_sucursal:= dsArticulosConsulta.FieldByName('id_sucursal').AsInteger;
            if Assigned(cdsArticulos.FindField('id_sucursal'+IntToStr(iId_sucursal))) then
            begin
              cdsArticulos.FieldByName('id_sucursal'+IntToStr(iId_sucursal)).AsInteger := dsArticulosConsulta.FieldByName('Id_sucursal').AsInteger;
              cdsArticulos.FieldByName('des_sucursal'+IntToStr(iId_sucursal)).AsString := dsArticulosConsulta.FieldByName('des_sucursal').AsString;
              cdsArticulos.FieldByName('stock_minimo'+IntToStr(iId_sucursal)).AsFloat := dsArticulosConsulta.FieldByName('stock_minimo').AsFloat;
              cdsArticulos.FieldByName('stock_maximo'+IntToStr(iId_sucursal)).AsFloat := dsArticulosConsulta.FieldByName('stock_maximo').AsFloat;
              cdsArticulos.FieldByName('id_nivel_rotacion'+IntToStr(iId_sucursal)).AsInteger := dsArticulosConsulta.FieldByName('id_nivel_rotacion').AsInteger;
              if dsArticulosConsulta.FieldByName('mym_excluye_calculo_automat').AsInteger=1 then
                cdsArticulos.FieldByName('excluir_calc_aut'+IntToStr(iId_sucursal)).AsString := 'SiExcluyeAutomatico'
              else
                cdsArticulos.FieldByName('excluir_calc_aut'+IntToStr(iId_sucursal)).AsString := '';
            end;
//
            dsArticulosConsulta.Next;
          end;

          // Para que agrege la ultima pasada
          if sId_Articulo<>'******************' then
            cdsArticulos.Post;
        end;
      end;
    finally
      //NADA
    end;

    ///////////////////////////////////////////////////////////////////////////////////////////////
    // A este punto llega el dataset listo para setear el stock
    ///////////////////////////////////////////////////////////////////////////////////////////////
    if cdsArticulos.RecordCount > 0 then
    begin
      try
        cdsArticulos.First;
        while not cdsArticulos.Eof do
        begin

          iDiasPeriodoAnalisis := 0;
          iDiasDeStock         := 0;
          dCoefStockMaximo     := 1;

          if (cdsArticulos.FieldByName('mym_dias_periodoAnalisis_suc').AsInteger > 0) then
          begin
            iDiasPeriodoAnalisis := cdsArticulos.FieldByName('mym_dias_periodoAnalisis_suc').AsInteger;
          end;

          if iDiasPeriodoAnalisis=0 then
          begin
            if (cdsArticulos.FieldByName('mym_dias_periodoanalisis_art').AsInteger > 0) then
            begin
              iDiasPeriodoAnalisis := cdsArticulos.FieldByName('mym_dias_periodoanalisis_art').AsInteger;
            end;
          end;

          if iDiasPeriodoAnalisis>0 then
          begin

            dtFechaHasta := DateOf(Now-1);
            dtFechaDesde := DateOf(dtFechaHasta-iDiasPeriodoAnalisis);

            sFecha_desde := FormatDateTime('yyyy/mm/dd',dtFechaDesde);
            sFecha_hasta := FormatDateTime('yyyy/mm/dd',dtFechaHasta);

            if Assigned(dsSucursales) then
            begin
              dsSucursales.First;
              dsSucursales.Filter := 'habilita_calc_minmax_autom = 1';
              dsSucursales.Filtered := True;

              if dsSucursales.RecordCount > 0 then
              begin
                if dsSucursales.Active then
                begin
                  dsSucursales.First;
                  while not dsSucursales.Eof do
                  begin
                    if cdsArticulos.FieldByName('excluir_calc_aut'+IntToStr(iId_sucursal)).AsString='' then
                    begin
                      iId_sucursal:= dsSucursales.FieldByName('id_sucursal').AsInteger;

                      oNivel_de_Rotacion:= _ObtenerNivele_Rotacion(pId_Empresa, iId_sucursal, cdsArticulos.FieldByName('id_articulo').AsString, dtFechaDesde, dtFechaHasta);
                      if Assigned(oNivel_de_Rotacion) then
                      begin

                        if Assigned(cdsArticulos.FindField('id_sucursal'+IntToStr(iId_sucursal))) then
                        begin
                          cdsArticulos.Edit;
                          cdsArticulos.FieldByName('id_nivel_rotacion'+IntToStr(iId_sucursal)).AsInteger := oNivel_de_Rotacion.Id_Nivel_Rotacion;

                          dsStock_min_sugerido := _Gestor_Articulos._Obtener_Stock_min_sugerido(pId_Empresa, sFecha_desde, sFecha_hasta, cdsArticulos.FieldByName('id_articulo').AsString, oNivel_de_Rotacion.Dias_StockMin, iId_sucursal);
                          if Assigned(dsStock_min_sugerido) then
                          begin
                            try
                              if dsStock_min_sugerido.RecordCount > 0 then
                              begin
                                if dsStock_min_sugerido.Active then
                                begin
                                  dsStock_min_sugerido.First;
                                  while not dsStock_min_sugerido.Eof do
                                  begin
                                    if Assigned(cdsArticulos.FindField('stock_minimo'+IntToStr(iId_sucursal))) then
                                    begin
                                      cdsArticulos.FieldByName('stock_minimo'+IntToStr(iId_sucursal)).AsFloat := RoundTo(dsStock_min_sugerido.FieldByName('stock_min_sugerido').AsFloat,0);
                                      if oNivel_de_Rotacion.Dias_StockMin>0 then
                                        dCoefStockMaximo := oNivel_de_Rotacion.Dias_StockMax/oNivel_de_Rotacion.Dias_StockMin;
                                      if dCoefStockMaximo<1 then
                                        dCoefStockMaximo:=1;
                                      cdsArticulos.FieldByName('stock_maximo'+IntToStr(iId_sucursal)).AsFloat := RoundTo(dsStock_min_sugerido.FieldByName('stock_min_sugerido').AsFloat * dCoefStockMaximo, 0);
                                    end;
                                    dsStock_min_sugerido.Next;
                                  end; // while not dsStock_min_sugerido.Eof
                                end; // if dsStock_min_sugerido.Active
                              end; // if dsStock_min_sugerido.RecordCount
                            finally
                              // nada
                            end;
                          end; // if Assigned(cdsArticulosConsulta)
                          cdsArticulos.Post;
                        end;
                        FreeAndNil(oNivel_de_Rotacion);
                      end
                      else // NO TIENE DEFINIDO NIVEL DE ROTACION USO LOS VALORES MANUALES
                      begin
                        if Assigned(cdsArticulos.FindField('id_sucursal'+IntToStr(iId_sucursal))) then
                        begin
                          cdsArticulos.Edit;
                          cdsArticulos.FieldByName('id_nivel_rotacion'+IntToStr(iId_sucursal)).AsInteger := 0;
                          {
                          // No hago nada porque no tengo el dato  Dias_StockMin
                          dsStock_min_sugerido := _Gestor_Articulos._Obtener_Stock_min_sugerido(pId_Empresa, sFecha_desde, sFecha_hasta, cdsArticulos.FieldByName('id_articulo').AsString, oNivel_de_Rotacion.Dias_StockMin, iId_sucursal);
                          if Assigned(dsStock_min_sugerido) then
                          begin
                            try
                              if dsStock_min_sugerido.RecordCount > 0 then
                              begin
                                if dsStock_min_sugerido.Active then
                                begin
                                  dsStock_min_sugerido.First;
                                  while not dsStock_min_sugerido.Eof do
                                  begin
                                    if Assigned(cdsArticulos.FindField('stock_minimo'+IntToStr(iId_sucursal))) then
                                    begin
                                      cdsArticulos.FieldByName('stock_minimo'+IntToStr(iId_sucursal)).AsFloat := RoundTo(dsStock_min_sugerido.FieldByName('stock_min_sugerido').AsFloat,0);
                                      if oNivel_de_Rotacion.Dias_StockMin>0 then
                                        cdsArticulos.FieldByName('stock_maximo'+IntToStr(iId_sucursal)).AsFloat := RoundTo(dsStock_min_sugerido.FieldByName('stock_min_sugerido').AsFloat * RoundTo(oNivel_de_Rotacion.Dias_StockMax/oNivel_de_Rotacion.Dias_StockMin, -4), 0);
                                    end;
                                    dsStock_min_sugerido.Next;
                                  end; // while not dsStock_min_sugerido.Eof
                                end; // if dsStock_min_sugerido.Active
                              end; // if dsStock_min_sugerido.RecordCount
                            finally
                              // nada
                            end;
                          end; // if Assigned(cdsArticulosConsulta)
                          }
                          cdsArticulos.Post;
                        end;
                      end; //if Assigned(oNivel_de_Rotacion) then
                    end;  //if cdsArticulos.FieldByName('excluir_calc_aut'+IntToStr(iId_sucursal)).AsInteger=0 then
                    dsSucursales.Next;
                  end;
                end;
              end;
            end;
          end;


          cdsArticulos.Next;
        end;  // while not cdsArticulos.Eof

        //Habiendo seteado el stock, guardo los cambios
        cdsArticulos.First;
        try
          sFiltro_Sucursales := '';
          if Assigned(dsSucursales) then
          begin
            if dsSucursales.RecordCount > 0 then
            begin
              dsSucursales.First;
              while not(dsSucursales.Eof) do
              begin
                if dsSucursales.FieldByName('habilita_calc_minmax_autom').AsInteger > 0 then
                begin
                  if Length(sFiltro_Sucursales) > 0 then
                    sFiltro_Sucursales:= sFiltro_Sucursales + ' OR ';
                  sFiltro_Sucursales:= sFiltro_Sucursales + 'id_sucursal='+dsSucursales.FieldByName('id_sucursal').AsString;
                end;
                dsSucursales.Next;
              end;
            end;
          end;

          oBarraProgreso := Nil;
          sMensaje_error := _Gestor_Articulos._Modificar_Articulos_Stock_x_Sucursal(50, cdsArticulos, oBarraProgreso, pId_empresa, sFiltro_Sucursales);
          if sMensaje_error <> EmptyStr then
          begin
            _Grabar_LogErrores_Comunicacion(-7, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - '+'No se pudo Modificar Stock por Sucursal por error: '+sMensaje_error+Chr(13)+'Proceso de actualización INCOMPLETO!!!');
            Exit;
          end;

          _Grabar_LogErrores_Comunicacion(-7, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - '+'Proceso finalizado exitosamente');

        except
          on EError_Guardar : Exception do
          begin
            _Grabar_LogErrores_Comunicacion(-7, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ EError_Guardar.Message);
          end;
        end;  // try

      finally
        //
      end;
    end;
  end;

end;

function TVSComunicacionParaReplicacion01._Actualizar_Precios_Programados: String;
var
  dsLista_Articulos_Consulta : TDataSet;
  sWhereFiltro : String;
  oArticulo    : TArticulo;
  dPcio_com_siva_bruto_Ant   : Double;
  dPcio_vta_siva_Ant         : Double;
  sFormulario, sMensajeError : String;
  iCantidad_Ok: Integer;
  iCantidad_Error: Integer;
begin
  Result:='';
  iCantidad_Ok:=0;
  iCantidad_Error:=0;
  sWhereFiltro  := 'WHERE pcio_vta_siva_programado <> 0';
  dsLista_Articulos_Consulta:= _Gestor_Articulos._BuscarArticulos_SeleccionMultimple_ActuPreciosImportarXLS(sWhereFiltro);
  if Assigned(dsLista_Articulos_Consulta) then
  begin
    dsLista_Articulos_Consulta.First;
    while not dsLista_Articulos_Consulta.Eof do
    begin
      if Assigned(dsLista_Articulos_Consulta.FindField('id_articulo')) then
      begin
        if Length(dsLista_Articulos_Consulta.FieldByName('id_articulo').AsString) > 0 then
        begin
          oArticulo := _Gestor_Articulos._Buscar(dsLista_Articulos_Consulta.FieldByName('id_articulo').AsString);
          if Assigned(oArticulo) then
          begin

            if oArticulo.Modalidad_relacion_pc_pv='PC' then
              dPcio_com_siva_bruto_Ant := oArticulo.Pcio_com_siva_bruto
            else if oArticulo.Modalidad_relacion_pc_pv='PR' then
              dPcio_com_siva_bruto_Ant := oArticulo.Pcio_com_siva_rep;

            dPcio_vta_siva_Ant       := oArticulo.Pcio_vta_siva;
            if oArticulo.Pcio_com_siva_bruto_programado > 0 then
            begin
              if oArticulo.Modalidad_relacion_pc_pv='PC' then
              begin
                oArticulo.Pcio_com_siva_bruto := oArticulo.Pcio_com_siva_bruto_programado;
                oArticulo.Pcio_vta_siva       := oArticulo.Pcio_vta_siva_programado;
                oArticulo.Pcio_com_siva       := oArticulo.Pcio_com_siva_bruto - oArticulo.Pcio_com_siva_bruto * (oArticulo.Porc_desc_pcio_com_bruto/100) + oArticulo.Importe_otro;
                oArticulo.Coeficiente_ganancia     := oArticulo.Pcio_vta_siva / oArticulo.Pcio_com_siva;
              end
              else if oArticulo.Modalidad_relacion_pc_pv='PR' then
              begin
                oArticulo.Pcio_com_siva_rep   := oArticulo.Pcio_com_siva_bruto_programado;
                oArticulo.Pcio_vta_siva       := oArticulo.Pcio_vta_siva_programado;
                //oArticulo.Pcio_com_siva       := oArticulo.Pcio_com_siva_rep - oArticulo.Pcio_com_siva_rep * (oArticulo.Porc_desc_pcio_com_bruto/100) + oArticulo.Importe_otro;
                oArticulo.Coeficiente_ganancia     := oArticulo.Pcio_vta_siva / oArticulo.Pcio_com_siva_rep;
              end;
            end;
            oArticulo.Pcio_com_siva_bruto_programado:=0;
            oArticulo.Pcio_vta_siva_programado      :=0;
            //sFormulario:= 'Actualización de precios programada';
            sFormulario:= 'Act.Prog. Origen: '+oArticulo.Modulo_Origen_programado;
            oArticulo.Modulo_Origen_programado:= '';

            //' - Modalidad: Programada'
            //' - Modalidad: Instantánea'

            sMensajeError := Trim(_Gestor_Articulos._Modificar(oArticulo, oArticulo.Id_articulo, dPcio_com_siva_bruto_Ant, dPcio_vta_siva_Ant, oArticulo.Pcio_com_siva_bruto, oArticulo.Pcio_com_siva_rep, oParametro.Id_Empresa, oParametro.Id_Sucursal, 'Zafiro', sFormulario, ''));
            if Length(sMensajeError)>0 then
            begin
              sMensajeError:= ' - Error al modificar el Articulo con Id: '+oArticulo.Id_articulo + ' - ' + oArticulo.Des_articulo +', tiene problemas al guardar (modificar).' + ' - Error: ' + sMensajeError;
              _Grabar_LogErrores_Comunicacion(-4, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ sMensajeError);
              iCantidad_Error:=  iCantidad_Error + 1;
            end
            else
            begin
              iCantidad_Ok:=  iCantidad_Ok + 1;
            end;
            FreeAndNil(oArticulo);
          end;
        end;
      end;
      dsLista_Articulos_Consulta.Next;
    end;
    _Grabar_LogErrores_Comunicacion(-4, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Se actualizaron '+IntToStr(iCantidad_Ok)+' correctamente. Ocurrieron ' + IntToStr(iCantidad_Error)+' errores al actualizar.');
  end;
end;

function TVSComunicacionParaReplicacion01._Actualizar_Zafiro_BI(pParametro_Zafiro_BI: TParametro_Zafiro_BI; pFecha_Hasta: TDate): String;
var
  oGestor_Lanzador_Zafiro_Local: TGestor_Lanzador_Zafiro;
  oGestor_Comprobante_Ventas : TGestor_Comprobante_Ventas;
  dtFecha                    : TDate;
  dtFecha_Comp_Ingreso       : TDate;
  iCantidad_Dias_Vtas_Ant    : Integer;
  iCantidad_Meses_Hacia_Atras: Integer;
  dtFecha_Stock_Por_Mes      : TDate;
  dsComprobantes             : TDataSet;
  dsArticulosAltas           : TDataSet;
  dsArticulosModificaciones  : TDataSet;

  dsClientesAltas            : TDataSet;
  dsClientesModificaciones   : TDataSet;
  dsPlanesECAltas            : TDataSet;
  dsPlanesECModificaciones   : TDataSet;

  dsEntidadConvenioAltas     : TDataSet;
  dsEntidadConvenioModificaciones : TDataSet;
  dsPlanesEDAltas            : TDataSet;
  dsPlanesEDModificaciones   : TDataSet;

  dsCajasXSucAltas           : TDataSet;
  dsCajasXSucModificaciones  : TDataSet;

  dsProveedoresAltas            : TDataSet;
  dsProveedoresModificaciones   : TDataSet;

  oFactura_ND_Presupuesto    : TFactura_ND_Presupuesto;
  oNota_de_Credito           : TNota_de_Credito;
  oLista_Comprobantes_Sin_Subir: TObjectList;
  sResultado                 : String;
  bError                     : Boolean;
  bHayComprobantes           : Boolean;
  iLote                      : Integer;
  dtFechaHora_Ini, dtFechaHora_Fin: TDateTime;
  oGestor_Vendedores         : TGestor_Vendedores;
  oGestor_Clientes           : TGestor_Clientes;
  oGestor_Proveedores        : TGestor_Proveedores;
  iItem: SmallInt;
  //Formato: TFormatSettings;
  sRespuesta : String;
  dsLin_Comprobante         : TDataSet;
  iCant_Lineas              : Integer;

  slConfig_BI_externo: TStringList;
  iLinea : SmallInt;
  sCadena: String;

  dsStock_Por_Mes           : TDataSet;
  dtFecha_Vtas_Desde        : TDate;
  dtFecha_Vtas_Hasta        : TDate;

  oGestor_Comprobante_Stock : TGestor_Comprobante_Stock;
  dsComprobantes_Ingresos   : TDataSet;
  dsLin_Comprobante_Ingreso : TDataSet;
  oLista_Comprobantes_Ingresos_Sin_Subir: TObjectList;
  oIngreso_Mercaderia       : TIngreso_Mercaderia;
begin
  // Funcion que envia el oParametro y recupera todos los comprobantes que
  // no fueron informados.
  Result:= '';

  _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Inicializando proceso... ');

  oGestor_Lanzador_Zafiro_Local:= TGestor_Lanzador_Zafiro.Create(pParametro_Zafiro_BI.Cadena_Conexion_BD);
  oGestor_Vendedores := TGestor_Vendedores.Create;
  oGestor_Clientes   := TGestor_Clientes.Create(oParametro);
  oGestor_Proveedores:= TGestor_Proveedores.Create;

  // Articulos
  try
    dsArticulosAltas := oGestor_Lanzador_Zafiro_Local._Buscar_Articulos_Altas_Sin_Subir_Zafiro_BI;
  except
  end;

  if Assigned(dsArticulosAltas) then
  begin
    if dsArticulosAltas.RecordCount > 0 then
    begin
      _Grabar_LogErrores_Comunicacion(-5, '   Registrando '+ IntToStr(dsArticulosAltas.RecordCount)+' Altas de Articulos en Zafiro BI - '+FormatDateTime('hh:mm:ss', Now));
      if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
      begin
        _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Reconexión con la BD Zafiro BI...');
        if Assigned(oGestor_Lanzador_Zafiro_Local) then
          FreeAndNil(oGestor_Lanzador_Zafiro_Local);
        oGestor_Lanzador_Zafiro_Local:= TGestor_Lanzador_Zafiro.Create(pParametro_Zafiro_BI.Cadena_Conexion_BD);
        if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
          _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - No se pudo establecer conexión con la BD Zafiro BI. (5)');
      end;

      if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI then
      begin
        sResultado := oGestor_Lanzador_Zafiro_Local._Subir_Articulos_Altas_Zafiro_BI(pParametro_Zafiro_BI.Vs_Id_Cliente, dsArticulosAltas);
        if Length(Trim(sResultado))>0 then
        begin
          _Grabar_LogErrores_Comunicacion(-5, '   '+sResultado);
        end
        else
        begin
          _Grabar_LogErrores_Comunicacion(-5, '      Se Registraron '+ IntToStr(dsArticulosAltas.RecordCount)+' Articulos en Zafiro BI - '+FormatDateTime('hh:mm:ss', Now));
        end;
      end;
    end;
  end;

  try
    dsArticulosModificaciones := oGestor_Lanzador_Zafiro_Local._Buscar_Articulos_Modificaciones_Sin_Subir_Zafiro_BI;
  except
  end;

  if Assigned(dsArticulosModificaciones) then
  begin
    if dsArticulosModificaciones.RecordCount > 0 then
    begin
      _Grabar_LogErrores_Comunicacion(-5, '   Registrando '+ IntToStr(dsArticulosModificaciones.RecordCount)+' Modificaciones de Articulos en Zafiro BI - '+FormatDateTime('hh:mm:ss', Now));
      if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
      begin
        _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Reconexión con la BD Zafiro BI...');
        if Assigned(oGestor_Lanzador_Zafiro_Local) then
          FreeAndNil(oGestor_Lanzador_Zafiro_Local);
        oGestor_Lanzador_Zafiro_Local:= TGestor_Lanzador_Zafiro.Create(pParametro_Zafiro_BI.Cadena_Conexion_BD);
        if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
          _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - No se pudo establecer conexión con la BD Zafiro BI. (6)');
      end;

      if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI then
      begin
        sResultado := oGestor_Lanzador_Zafiro_Local._Subir_Articulos_Modificaciones_Zafiro_BI(pParametro_Zafiro_BI.Vs_Id_Cliente, dsArticulosModificaciones);
        if Length(Trim(sResultado))>0 then
        begin
          _Grabar_LogErrores_Comunicacion(-5, '   '+sResultado);
        end
        else
        begin
          _Grabar_LogErrores_Comunicacion(-5, '      Se Registraron '+ IntToStr(dsArticulosModificaciones.RecordCount)+' Modificaciones de Articulos en Zafiro BI - '+FormatDateTime('hh:mm:ss', Now));
        end;
      end;
    end;
  end;

  LiberarMemoria;

  // Clientes Altas
  try
    dsClientesAltas := oGestor_Lanzador_Zafiro_Local._Buscar_Clientes_Altas_Sin_Subir_Zafiro_BI;
  except
  end;

  try
    dsPlanesECAltas := oGestor_Lanzador_Zafiro_Local._Buscar_Clientes_Planes_Altas_Sin_Subir_Zafiro_BI;
  except
  end;

  if Assigned(dsClientesAltas) then
  begin
    with dsClientesAltas do
    begin
      if RecordCount > 0 then
      begin
        _Grabar_LogErrores_Comunicacion(-5, '   Registrando '+ IntToStr(RecordCount)+' Altas de Clientes en Zafiro BI - '+FormatDateTime('hh:mm:ss', Now));
        if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
        begin
          _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Reconexión con la BD Zafiro BI...');
          if Assigned(oGestor_Lanzador_Zafiro_Local) then
            FreeAndNil(oGestor_Lanzador_Zafiro_Local);
          oGestor_Lanzador_Zafiro_Local:= TGestor_Lanzador_Zafiro.Create(pParametro_Zafiro_BI.Cadena_Conexion_BD);
          if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
            _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - No se pudo establecer conexión con la BD Zafiro BI. (5)');
        end;

        if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI then
        begin
          sResultado := oGestor_Lanzador_Zafiro_Local._Subir_Clientes_Altas_Zafiro_BI(pParametro_Zafiro_BI.Vs_Id_Cliente, dsClientesAltas, dsPlanesECAltas);
          if Length(Trim(sResultado))>0 then
          begin
            _Grabar_LogErrores_Comunicacion(-5, '   '+sResultado);
          end
          else
          begin
            _Grabar_LogErrores_Comunicacion(-5, '      Se Registraron '+ IntToStr(RecordCount)+' Clientes en Zafiro BI - '+FormatDateTime('hh:mm:ss', Now));
          end;
        end;
      end;
    end;
  end;

  // Clientes Modificaciones
  try
    dsClientesModificaciones := oGestor_Lanzador_Zafiro_Local._Buscar_Clientes_Modificaciones_Sin_Subir_Zafiro_BI;
  except
  end;

  try
    dsPlanesECModificaciones := oGestor_Lanzador_Zafiro_Local._Buscar_Clientes_Planes_Modificaciones_Sin_Subir_Zafiro_BI;
  except
  end;

  if Assigned(dsClientesModificaciones) then
  begin
    with dsClientesModificaciones do
    begin
      if RecordCount > 0 then
      begin
        _Grabar_LogErrores_Comunicacion(-5, '   Registrando '+ IntToStr(RecordCount)+' Modificaciones de Clientes en Zafiro BI - '+FormatDateTime('hh:mm:ss', Now));
        if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
        begin
          _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Reconexión con la BD Zafiro BI...');
          if Assigned(oGestor_Lanzador_Zafiro_Local) then
            FreeAndNil(oGestor_Lanzador_Zafiro_Local);
          oGestor_Lanzador_Zafiro_Local:= TGestor_Lanzador_Zafiro.Create(pParametro_Zafiro_BI.Cadena_Conexion_BD);
          if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
            _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - No se pudo establecer conexión con la BD Zafiro BI. (5)');
        end;

        if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI then
        begin
          sResultado := oGestor_Lanzador_Zafiro_Local._Subir_Clientes_Modificaciones_Zafiro_BI(pParametro_Zafiro_BI.Vs_Id_Cliente, dsClientesModificaciones, dsPlanesECModificaciones);
          if Length(Trim(sResultado))>0 then
          begin
            _Grabar_LogErrores_Comunicacion(-5, '   '+sResultado);
          end
          else
          begin
            _Grabar_LogErrores_Comunicacion(-5, '      Se Registraron '+ IntToStr(RecordCount)+' Modificaciones de Clientes en Zafiro BI - '+FormatDateTime('hh:mm:ss', Now));
          end;
        end;
      end;
    end;
  end;

  LiberarMemoria;

  // Entidad Debito Altas
  try
    dsEntidadConvenioAltas := oGestor_Lanzador_Zafiro_Local._Buscar_ED_Altas_Sin_Subir_Zafiro_BI;
  except
  end;

  try
    dsPlanesEDAltas := oGestor_Lanzador_Zafiro_Local._Buscar_ED_Planes_Altas_Sin_Subir_Zafiro_BI;
  except
  end;

  if Assigned(dsEntidadConvenioAltas) then
  begin
    with dsEntidadConvenioAltas do
    begin
      if RecordCount > 0 then
      begin
        _Grabar_LogErrores_Comunicacion(-5, '   Registrando '+ IntToStr(RecordCount)+' Altas de Entidad Debito en Zafiro BI - '+FormatDateTime('hh:mm:ss', Now));
        if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
        begin
          _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Reconexión con la BD Zafiro BI...');
          if Assigned(oGestor_Lanzador_Zafiro_Local) then
            FreeAndNil(oGestor_Lanzador_Zafiro_Local);
          oGestor_Lanzador_Zafiro_Local:= TGestor_Lanzador_Zafiro.Create(pParametro_Zafiro_BI.Cadena_Conexion_BD);
          if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
            _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - No se pudo establecer conexión con la BD Zafiro BI. (5)');
        end;

        if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI then
        begin
          sResultado := oGestor_Lanzador_Zafiro_Local._Subir_ED_Altas_Zafiro_BI(pParametro_Zafiro_BI.Vs_Id_Cliente, dsEntidadConvenioAltas, dsPlanesEDAltas);
          if Length(Trim(sResultado))>0 then
          begin
            _Grabar_LogErrores_Comunicacion(-5, '   '+sResultado);
          end
          else
          begin
            _Grabar_LogErrores_Comunicacion(-5, '      Se Registraron '+ IntToStr(RecordCount)+' Entidad Debito en Zafiro BI - '+FormatDateTime('hh:mm:ss', Now));
          end;
        end;
      end;
    end;
  end;

  // Entidad Debito Modificaciones
  try
    dsEntidadConvenioModificaciones := oGestor_Lanzador_Zafiro_Local._Buscar_ED_Modificaciones_Sin_Subir_Zafiro_BI;
  except
  end;

  try
    dsPlanesEDModificaciones := oGestor_Lanzador_Zafiro_Local._Buscar_ED_Planes_Modificaciones_Sin_Subir_Zafiro_BI;
  except
  end;

  if Assigned(dsEntidadConvenioModificaciones) then
  begin
    with dsEntidadConvenioModificaciones do
    begin
      if RecordCount > 0 then
      begin
        _Grabar_LogErrores_Comunicacion(-5, '   Registrando '+ IntToStr(RecordCount)+' Modificaciones de Entidad Debito en Zafiro BI - '+FormatDateTime('hh:mm:ss', Now));
        if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
        begin
          _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Reconexión con la BD Zafiro BI...');
          if Assigned(oGestor_Lanzador_Zafiro_Local) then
            FreeAndNil(oGestor_Lanzador_Zafiro_Local);
          oGestor_Lanzador_Zafiro_Local:= TGestor_Lanzador_Zafiro.Create(pParametro_Zafiro_BI.Cadena_Conexion_BD);
          if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
            _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - No se pudo establecer conexión con la BD Zafiro BI. (5)');
        end;

        if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI then
        begin
          sResultado := oGestor_Lanzador_Zafiro_Local._Subir_ED_Modificaciones_Zafiro_BI(pParametro_Zafiro_BI.Vs_Id_Cliente, dsEntidadConvenioModificaciones, dsPlanesEDModificaciones);
          if Length(Trim(sResultado))>0 then
          begin
            _Grabar_LogErrores_Comunicacion(-5, '   '+sResultado);
          end
          else
          begin
            _Grabar_LogErrores_Comunicacion(-5, '      Se Registraron '+ IntToStr(RecordCount)+' Modificaciones de Entidad Debito en Zafiro BI - '+FormatDateTime('hh:mm:ss', Now));
          end;
        end;
      end;
    end;
  end;

  LiberarMemoria;

  // Cajas x Sucursal Altas
  try
    dsCajasXSucAltas := oGestor_Lanzador_Zafiro_Local._Buscar_Cajas_Altas_Sin_Subir_Zafiro_BI;
  except
  end;

  if Assigned(dsCajasXSucAltas) then
  begin
    with dsCajasXSucAltas do
    begin
      if RecordCount > 0 then
      begin
        _Grabar_LogErrores_Comunicacion(-5, '   Registrando '+ IntToStr(RecordCount)+' Altas de Cajas por Sucursal en Zafiro BI - '+FormatDateTime('hh:mm:ss', Now));
        if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
        begin
          _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Reconexión con la BD Zafiro BI...');
          if Assigned(oGestor_Lanzador_Zafiro_Local) then
            FreeAndNil(oGestor_Lanzador_Zafiro_Local);
          oGestor_Lanzador_Zafiro_Local:= TGestor_Lanzador_Zafiro.Create(pParametro_Zafiro_BI.Cadena_Conexion_BD);
          if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
            _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - No se pudo establecer conexión con la BD Zafiro BI. (5)');
        end;

        if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI then
        begin
          sResultado := oGestor_Lanzador_Zafiro_Local._Subir_Cajas_Altas_Zafiro_BI(pParametro_Zafiro_BI.Vs_Id_Cliente, dsCajasXSucAltas);
          if Length(Trim(sResultado))>0 then
          begin
            _Grabar_LogErrores_Comunicacion(-5, '   '+sResultado);
          end
          else
          begin
            _Grabar_LogErrores_Comunicacion(-5, '      Se Registraron '+ IntToStr(RecordCount)+' Cajas por Sucursal en Zafiro BI - '+FormatDateTime('hh:mm:ss', Now));
          end;
        end;
      end;
    end;
  end;
  // FIN. Cajas x Sucursal Altas

  // Cajas x Sucursal Modificaciones
  try
    dsCajasXSucModificaciones := oGestor_Lanzador_Zafiro_Local._Buscar_Cajas_Modificaciones_Sin_Subir_Zafiro_BI;
  except

  end;

  if Assigned(dsCajasXSucModificaciones) then
  begin
    with dsCajasXSucModificaciones do
    begin
      if RecordCount > 0 then
      begin
        _Grabar_LogErrores_Comunicacion(-5, '   Registrando '+ IntToStr(RecordCount)+' Modificaciones de Cajas en Zafiro BI - '+FormatDateTime('hh:mm:ss', Now));
        if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
        begin
          _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Reconexión con la BD Zafiro BI...');
          if Assigned(oGestor_Lanzador_Zafiro_Local) then
            FreeAndNil(oGestor_Lanzador_Zafiro_Local);
          oGestor_Lanzador_Zafiro_Local:= TGestor_Lanzador_Zafiro.Create(pParametro_Zafiro_BI.Cadena_Conexion_BD);
          if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
            _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - No se pudo establecer conexión con la BD Zafiro BI. (5)');
        end;

        if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI then
        begin
          sResultado := oGestor_Lanzador_Zafiro_Local._Subir_Cajas_Modificaciones_Zafiro_BI(pParametro_Zafiro_BI.Vs_Id_Cliente, dsCajasXSucModificaciones);
          if Length(Trim(sResultado))>0 then
          begin
            _Grabar_LogErrores_Comunicacion(-5, '   '+sResultado);
          end
          else
          begin
            _Grabar_LogErrores_Comunicacion(-5, '      Se Registraron '+ IntToStr(RecordCount)+' Modificaciones de Cajas en Zafiro BI - '+FormatDateTime('hh:mm:ss', Now));
          end;
        end;
      end;
    end;
  end;

  // FIN. Cajas x Sucursal Modificaciones

  LiberarMemoria;

  // Proveedores Altas
  try
    dsProveedoresAltas := oGestor_Lanzador_Zafiro_Local._Buscar_Proveedores_Altas_Sin_Subir_Zafiro_BI;
  except
  end;

  if Assigned(dsProveedoresAltas) then
  begin
    with dsProveedoresAltas do
    begin
      if RecordCount > 0 then
      begin
        _Grabar_LogErrores_Comunicacion(-5, '   Registrando '+ IntToStr(RecordCount)+' Altas de Proveedores en Zafiro BI - '+FormatDateTime('hh:mm:ss', Now));
        if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
        begin
          _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Reconexión con la BD Zafiro BI...');
          if Assigned(oGestor_Lanzador_Zafiro_Local) then
            FreeAndNil(oGestor_Lanzador_Zafiro_Local);
          oGestor_Lanzador_Zafiro_Local:= TGestor_Lanzador_Zafiro.Create(pParametro_Zafiro_BI.Cadena_Conexion_BD);
          if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
            _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - No se pudo establecer conexión con la BD Zafiro BI. (5)');
        end;

        if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI then
        begin
          sResultado := oGestor_Lanzador_Zafiro_Local._Subir_Proveedores_Altas_Zafiro_BI(pParametro_Zafiro_BI.Vs_Id_Cliente, dsProveedoresAltas);
          if Length(Trim(sResultado))>0 then
          begin
            _Grabar_LogErrores_Comunicacion(-5, '   '+sResultado);
          end
          else
          begin
            _Grabar_LogErrores_Comunicacion(-5, '      Se Registraron '+ IntToStr(RecordCount)+' Proveedores en Zafiro BI - '+FormatDateTime('hh:mm:ss', Now));
          end;
        end;
      end;
    end;
  end;

  // Proveedores Modificaciones
  try
    dsProveedoresModificaciones := oGestor_Lanzador_Zafiro_Local._Buscar_Proveedores_Modificaciones_Sin_Subir_Zafiro_BI;
  except
  end;

  if Assigned(dsProveedoresModificaciones) then
  begin
    with dsProveedoresModificaciones do
    begin
      if RecordCount > 0 then
      begin
        _Grabar_LogErrores_Comunicacion(-5, '   Registrando '+ IntToStr(RecordCount)+' Modificaciones de Proveedores en Zafiro BI - '+FormatDateTime('hh:mm:ss', Now));
        if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
        begin
          _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Reconexión con la BD Zafiro BI...');
          if Assigned(oGestor_Lanzador_Zafiro_Local) then
            FreeAndNil(oGestor_Lanzador_Zafiro_Local);
          oGestor_Lanzador_Zafiro_Local:= TGestor_Lanzador_Zafiro.Create(pParametro_Zafiro_BI.Cadena_Conexion_BD);
          if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
            _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - No se pudo establecer conexión con la BD Zafiro BI. (5)');
        end;

        if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI then
        begin
          sResultado := oGestor_Lanzador_Zafiro_Local._Subir_Proveedores_Modificaciones_Zafiro_BI(pParametro_Zafiro_BI.Vs_Id_Cliente, dsProveedoresModificaciones);
          if Length(Trim(sResultado))>0 then
          begin
            _Grabar_LogErrores_Comunicacion(-5, '   '+sResultado);
          end
          else
          begin
            _Grabar_LogErrores_Comunicacion(-5, '      Se Registraron '+ IntToStr(RecordCount)+' Modificaciones de Proveedores en Zafiro BI - '+FormatDateTime('hh:mm:ss', Now));
          end;
        end;
      end;
    end;
  end;

  LiberarMemoria;

  dtFecha := pParametro_Zafiro_BI.Fecha_Ultima_Actualizacion + 1 - 40; // 40 dias hacia atras
  dtFecha_Comp_Ingreso := pParametro_Zafiro_BI.Fecha_Ultima_Actualizacion_Comp_Ingreso + 1 - 40; // 40 dias hacia atras
  iCantidad_Dias_Vtas_Ant := 120;
  iCantidad_Meses_Hacia_Atras := 6;

  // Leo el archivo config_BI.txt
  // Configuraciones
  if FileExists('c:\vs-comunicacion\config_BI.txt') then
  begin
    slConfig_BI_externo:= TStringList.Create;
    slConfig_BI_externo.LoadFromFile('c:\vs-comunicacion\config_BI.txt');

    //
    // Comprobantes Fecha Desde
    if slConfig_BI_externo.Count>0 then
    begin
      sCadena:= Copy(slConfig_BI_externo[0], 50, length(slConfig_BI_externo[0]) - 50 + 1);  // Comprobantes Fecha Desde                        :01/01/2023
    end;
    try
      dtFecha := StrToDate(Trim(sCadena));
      dtFecha_Comp_Ingreso := StrToDate(Trim(sCadena));
    except
      // Vuelvo a poner la configuracion original
      // dtFecha := pParametro_Zafiro_BI.Fecha_Ultima_Actualizacion + 1 - 40; // 40 dias hacia atras
    end;
    //
    //

    //
    // Calculo Stock x Mes. Cantidad de dias Ventas Anteriores
    if slConfig_BI_externo.Count>1 then
    begin
      sCadena:= Copy(slConfig_BI_externo[1], 50, length(slConfig_BI_externo[1]) - 50 + 1);  // Cantidad de dias Ventas Anteriores              :120
    end;
    try
      iCantidad_Dias_Vtas_Ant := StrToInt(Trim(sCadena));
    except
      // iCantidad_Dias_Vtas_Ant := 120; // 120 dias hacia atras
    end;
    //
    //

    //
    // Calculo Stock x Mes. Cantidad Meses Hacia Atras :6
    if slConfig_BI_externo.Count>2 then
    begin
      sCadena:= Copy(slConfig_BI_externo[2], 50, length(slConfig_BI_externo[2]) - 50 + 1);  // Calculo Stock x Mes. Cantidad Meses Hacia Atras :6
    end;
    try
      iCantidad_Meses_Hacia_Atras := StrToInt(Trim(sCadena));
    except
      //
    end;
    //
    //

    FreeAndNil(slConfig_BI_externo);
  end;
  //

  ///////////////////////////////////////////
  // Comprobantes de Venta                 //
  ///////////////////////////////////////////

  oGestor_Comprobante_Ventas := TGestor_Comprobante_Ventas.Create(oParametro);

  while dtFecha<=pFecha_Hasta do
  begin

    bHayComprobantes := True;
    iLote            := 1;
    while bHayComprobantes do
    begin

      LiberarMemoria;

      //if Assigned(oGestor_Lanzador_Zafiro_Local) then
      //  FreeAndNil(oGestor_Lanzador_Zafiro_Local);
      //oGestor_Lanzador_Zafiro_Local:= TGestor_Lanzador_Zafiro.Create;

      try
        dsComprobantes:= oGestor_Lanzador_Zafiro_Local._Buscar_Comprobantes_Sin_Subir_Zafiro_BI(FormatDateTime('yyyy/mm/dd', dtFecha), FormatDateTime('yyyy/mm/dd', dtFecha), 100);
      except
      end;


      if Assigned(dsComprobantes) then
      begin

        oLista_Comprobantes_Sin_Subir := TObjectList.Create(True);

        with dsComprobantes do
        begin
          First;
          if RecordCount > 0 then
          begin

            dtFechaHora_Ini:= Now;
            _Grabar_LogErrores_Comunicacion(-5, '   Fecha: '+ FormatDateTime('dd/mm/yyyy', dtFecha)+' Lote:'+IntToStr(iLote)+' - Registrando '+ IntToStr(RecordCount)+' Comprobantes en BI - '+FormatDateTime('hh:mm:ss', dtFechaHora_Ini));

            while not eof do
            begin

              if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
              begin
                _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Reconexión con la BD Zafiro BI...');
                if Assigned(oGestor_Lanzador_Zafiro_Local) then
                  FreeAndNil(oGestor_Lanzador_Zafiro_Local);
                oGestor_Lanzador_Zafiro_Local:= TGestor_Lanzador_Zafiro.Create(pParametro_Zafiro_BI.Cadena_Conexion_BD);
                if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
                  _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - No se pudo establecer conexión con la BD Zafiro BI. (7)');

                // Para que salga y vuelva a empezar
                Break;
                //
              end;

              if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI then
              begin

                if oGestor_Lanzador_Zafiro_Local._Buscar_Comprobantes_Zafiro_BI(pParametro_Zafiro_BI.Vs_Id_Cliente, FieldByName('id_empresa').AsInteger, FieldByName('id_comprobante').AsString)= False then
                begin
                  if FieldByName('tipo_comprobante').AsString ='NC' then
                  begin
                    try
                      oNota_de_Credito := oGestor_Comprobante_Ventas._Buscar_NC(FieldByName('id_empresa').AsInteger, FieldByName('id_comprobante').AsString, oParametro);
                    except
                      oNota_de_Credito:= Nil;
                    end;
                    if oNota_de_Credito <> nil then
                    begin
                      oLista_Comprobantes_Sin_Subir.Add(oNota_de_Credito);
                    end;
                  end
                  else
                  begin
                    try
                      oFactura_ND_Presupuesto := oGestor_Comprobante_Ventas._Buscar_FAC_ND_PRE(FieldByName('id_empresa').AsInteger, FieldByName('id_comprobante').AsString, oParametro);
                    except
                      oFactura_ND_Presupuesto:= Nil;
                    end;
                    if oFactura_ND_Presupuesto <> nil then
                    begin
                      oLista_Comprobantes_Sin_Subir.Add(oFactura_ND_Presupuesto);
                    end;
                  end;
                end
                else
                begin
                  // Controla la cantidad de Lineas del comprobante
                  dsLin_Comprobante:= oGestor_Lanzador_Zafiro_Local._Buscar_Lin_Comprobantes_Zafiro_BI(pParametro_Zafiro_BI.Vs_Id_Cliente, FieldByName('id_empresa').AsInteger, FieldByName('id_comprobante').AsString);
                  if Assigned(dsLin_Comprobante)then
                  begin
                    iCant_Lineas := oGestor_Lanzador_Zafiro_Local._Buscar_Lin_Comprobantes_Zafiro(FieldByName('id_empresa').AsInteger, FieldByName('id_comprobante').AsString);
                    if dsLin_Comprobante.RecordCount <> iCant_Lineas then
                    begin
                      if FieldByName('tipo_comprobante').AsString ='NC' then
                      begin
                        try
                          oNota_de_Credito := oGestor_Comprobante_Ventas._Buscar_NC(FieldByName('id_empresa').AsInteger, FieldByName('id_comprobante').AsString, oParametro);
                        except
                          oNota_de_Credito:= Nil;
                        end;
                        if oNota_de_Credito <> nil then
                        begin
                          // Marco para identificar que es por diferencia de linea.
                          oNota_de_Credito.Calc_Marca2 := True;
                          //
                          if Assigned(oNota_de_Credito.Lineas) then
                          begin
                            for iItem := 0 to oNota_de_Credito.Lineas.Count-1 do
                            begin
                              if not (dsLin_Comprobante.Locate('item', TLin_Comprobante_Ventas(oNota_de_Credito.Lineas[iItem]).Item,[])) then
                              begin
                                // Utilizo Calc_No_Aprobado para identifical el item que no se encuentra
                                TLin_Comprobante_Ventas(oNota_de_Credito.Lineas[iItem]).Calc_No_Aprobado := True;
                                //
                              end;
                            end;
                          end;
                          oLista_Comprobantes_Sin_Subir.Add(oNota_de_Credito);
                        end;
                      end
                      else
                      begin
                        try
                          oFactura_ND_Presupuesto := oGestor_Comprobante_Ventas._Buscar_FAC_ND_PRE(FieldByName('id_empresa').AsInteger, FieldByName('id_comprobante').AsString, oParametro);
                        except
                          oFactura_ND_Presupuesto:= Nil;
                        end;
                        if oFactura_ND_Presupuesto <> nil then
                        begin
                          // Marco para identificar que es por diferencia de linea.
                          oFactura_ND_Presupuesto.Calc_Marca2 := True;
                          //
                          if Assigned(oFactura_ND_Presupuesto.Lineas) then
                          begin
                            for iItem := 0 to oFactura_ND_Presupuesto.Lineas.Count-1 do
                            begin
                              if not (dsLin_Comprobante.Locate('item', TLin_Comprobante_Ventas(oFactura_ND_Presupuesto.Lineas[iItem]).Item,[])) then
                              begin
                                // Utilizo Calc_No_Aprobado para identifical el item que no se encuentra
                                TLin_Comprobante_Ventas(oFactura_ND_Presupuesto.Lineas[iItem]).Calc_No_Aprobado := True;
                                //
                              end;
                            end;
                          end;
                          oLista_Comprobantes_Sin_Subir.Add(oFactura_ND_Presupuesto);
                        end;
                      end;
                    end
                    else
                    begin
                      // El Comprobante tiene la cantidad de Lineas correcta
                      sRespuesta:= oGestor_Lanzador_Zafiro_Local._Marcar_Como_Registrado_En_BI(FieldByName('id_empresa').AsInteger, FieldByName('id_comprobante').AsString);
                      if Length(sRespuesta)>0 then
                      begin
                        _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - '+sRespuesta);
                      end;
                    end;
                  end;
                end;
              end;

              Next;
            end;
          end;
          //Close;
        end;

        if oLista_Comprobantes_Sin_Subir.Count > 0 then
        begin
          _Grabar_LogErrores_Comunicacion(-5, '      Fecha: '+ FormatDateTime('dd/mm/yyyy', dtFecha)+' Lote:'+IntToStr(iLote)+' - Se Prepararon '+ IntToStr(oLista_Comprobantes_Sin_Subir.Count)+' Comprobantes en BI - '+FormatDateTime('hh:mm:ss', Now)+' Duración: '+FormatDateTime('hh:mm:ss', (Now-dtFechaHora_Ini)));
          if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
          begin
            _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Reconexión con la BD Zafiro BI...');
            if Assigned(oGestor_Lanzador_Zafiro_Local) then
              FreeAndNil(oGestor_Lanzador_Zafiro_Local);
            oGestor_Lanzador_Zafiro_Local:= TGestor_Lanzador_Zafiro.Create(pParametro_Zafiro_BI.Cadena_Conexion_BD);
            if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
              _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - No se pudo establecer conexión con la BD Zafiro BI. (8)');
          end;

          if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI then
          begin
            bError:= False;
            sResultado := oGestor_Lanzador_Zafiro_Local._Subir_Comprobantes_Zafiro_BI(pParametro_Zafiro_BI.Vs_Id_Cliente, oLista_Comprobantes_Sin_Subir, oParametro, oGestor_Comprobante_Ventas, oGestor_Clientes, oGestor_Vendedores);

            if Length(Trim(sResultado))>0 then
            begin
              _Grabar_LogErrores_Comunicacion(-5, '   Fecha: '+ FormatDateTime('dd/mm/yyyy', dtFecha)+' - '+ sResultado);
              if Length(Result)>0 then
                Result:= Result + Chr(13);
              Result:= Result + sResultado;
              bError:= True;
            end
            else
            begin
              if bError=False then
              begin
                dtFechaHora_Fin:= Now;
                _Grabar_LogErrores_Comunicacion(-5, '      Fecha: '+ FormatDateTime('dd/mm/yyyy', dtFecha)+' Lote:'+IntToStr(iLote)+' - Se Registraron '+ IntToStr(oLista_Comprobantes_Sin_Subir.Count)+' Comprobantes en BI - '+FormatDateTime('hh:mm:ss', dtFechaHora_Fin)+' Duración Total: '+FormatDateTime('hh:mm:ss', (dtFechaHora_Fin-dtFechaHora_Ini)));
              end;
            end;
            iLote:= iLote + 1;

          end;
//          for iItem := 0 to oLista_Comprobantes_Sin_Subir.Count-1 do
//          begin
//            if Assigned(TComprobante_Ventas(oLista_Comprobantes_Sin_Subir[iItem])) then
//            begin
//              if TComprobante_Ventas(oLista_Comprobantes_Sin_Subir[iItem]).Tipo_Comprobante=NC__NotaDeCredito then
//              begin
//                oNota_de_Credito := TNota_de_Credito(oLista_Comprobantes_Sin_Subir[iItem]);
//                FreeAndNil(oNota_de_Credito);
//              end
//              else
//              begin
//                oFactura_ND_Presupuesto := TFactura_ND_Presupuesto(oLista_Comprobantes_Sin_Subir[iItem]);
//                FreeAndNil(oFactura_ND_Presupuesto);
//              end;
//            end;
//          end;


        end;

        // Destruye la lista de Comprobantes
        //oLista_Comprobantes_Sin_Subir.Free;
        FreeAndNil(oLista_Comprobantes_Sin_Subir);
        //
        //FreeandNil(dsComprobantes);
        //
        //if dsComprobantes.Active then
        //  dsComprobantes.Close;

        //LiberarMemoria;

//        // Para reiniciar cada 5 lotes
//        if iLote mod 5 = 0 then
//        begin
//          if Assigned(oGestor_Lanzador_Zafiro_Local) then
//            FreeAndNil(oGestor_Lanzador_Zafiro_Local);
//
//          if Assigned(oGestor_Comprobante_Ventas) then
//            FreeAndNil(oGestor_Comprobante_Ventas);
//
//          if Assigned(oGestor_Vendedores) then
//            FreeAndNil(oGestor_Vendedores);
//
//          if Assigned(oGestor_Clientes) then
//            FreeAndNil(oGestor_Clientes);
//          Result:= 'Reiniciar';
//          Exit;
//        end;
//        //

      end
      else
      begin
        bHayComprobantes:= False;
        iLote:= 1;
      end;

    end;

    if Length(Result)=0 then
    begin

      //if not Assigned(oGestor_Lanzador_Zafiro_Local) then
      //  oGestor_Lanzador_Zafiro_Local:= TGestor_Lanzador_Zafiro.Create;

      sResultado := oGestor_Lanzador_Zafiro_Local._Actualizar_Fecha_Ult_Act_BI(dtFecha);
      if Length(Trim(sResultado))>0 then
      begin
        _Grabar_LogErrores_Comunicacion(-5, '   Fecha: '+ FormatDateTime('dd/mm/yyyy', dtFecha)+' - '+ sResultado);
        if Length(Result)>0 then
          Result:= Result + Chr(13);
        Result:= Result + sResultado;
      end;
    end;

    dtFecha := dtFecha + 1;
  end; // while dtFecha<=pFecha_Hasta do


  if Assigned(oGestor_Comprobante_Ventas) then
    FreeAndNil(oGestor_Comprobante_Ventas);

  if Assigned(oGestor_Vendedores) then
    FreeAndNil(oGestor_Vendedores);

  if Assigned(oGestor_Clientes) then
    FreeAndNil(oGestor_Clientes);

  ///////////////////////////////////////////
  // Fin Comprobantes de Venta             //
  ///////////////////////////////////////////


  ///////////////////////////////////////////
  // Comprobantes de Ingresos              //
  ///////////////////////////////////////////

//  // Blanquea bi_comp_ing_stock_sin_procesar
//  sResultado := oGestor_Lanzador_Zafiro_Local._Eliminar_BI_comp_ing_stock_sin_procesar(pParametro_Zafiro_BI.Vs_Id_Cliente);
//  if Length(Trim(sResultado))>0 then
//  begin
//    _Grabar_LogErrores_Comunicacion(-5, '   '+sResultado);
//  end
//  else
//  begin
//    _Grabar_LogErrores_Comunicacion(-5, '      Se elimino registros de Comprobantes de Ingreso de Stock Sin Procesar en Zafiro BI - '+FormatDateTime('hh:mm:ss', Now));
//  end;
//  //

  oGestor_Comprobante_Stock := TGestor_Comprobante_Stock.Create(oParametro);

  while dtFecha_Comp_Ingreso<=pFecha_Hasta do
  begin

    LiberarMemoria;

    bHayComprobantes := True;
    iLote            := 1;
    while bHayComprobantes do
    begin

      try
        dsComprobantes_Ingresos:= oGestor_Lanzador_Zafiro_Local._Buscar_Comprobantes_Ingresos_Sin_Subir_Zafiro_BI(FormatDateTime('yyyy/mm/dd', dtFecha_Comp_Ingreso), FormatDateTime('yyyy/mm/dd', dtFecha_Comp_Ingreso), 100);
      except
      end;

      if Assigned(dsComprobantes_Ingresos) then
      begin

        oLista_Comprobantes_Ingresos_Sin_Subir := TObjectList.Create(True);

        with dsComprobantes_Ingresos do
        begin
          First;
          if RecordCount > 0 then
          begin

            dtFechaHora_Ini:= Now;
            _Grabar_LogErrores_Comunicacion(-5, '   Fecha: '+ FormatDateTime('dd/mm/yyyy', dtFecha_Comp_Ingreso)+' Lote:'+IntToStr(iLote)+' - Registrando '+ IntToStr(RecordCount)+' Comprobantes de Ingresos en BI - '+FormatDateTime('hh:mm:ss', dtFechaHora_Ini));

            while not eof do
            begin

              if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
              begin
                _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Reconexión con la BD Zafiro BI...');
                if Assigned(oGestor_Lanzador_Zafiro_Local) then
                  FreeAndNil(oGestor_Lanzador_Zafiro_Local);
                oGestor_Lanzador_Zafiro_Local:= TGestor_Lanzador_Zafiro.Create(pParametro_Zafiro_BI.Cadena_Conexion_BD);
                if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
                  _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - No se pudo establecer conexión con la BD Zafiro BI. (7)');

                // Para que salga y vuelva a empezar
                Break;
                //
              end;

              if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI then
              begin

                if oGestor_Lanzador_Zafiro_Local._Buscar_Comprobantes_Ingresos_Zafiro_BI(pParametro_Zafiro_BI.Vs_Id_Cliente, FieldByName('id_empresa').AsInteger, FieldByName('id_comprobante').AsString)= False then
                begin
                  try
                    oIngreso_Mercaderia := oGestor_Comprobante_Stock._Buscar_Ingreso_Mercaderia(FieldByName('id_empresa').AsInteger, FieldByName('id_comprobante').AsString);
                  except
                    oIngreso_Mercaderia := Nil;
                  end;
                  if oIngreso_Mercaderia <> nil then
                  begin
                    if oIngreso_Mercaderia._Cantidad_Excluidos_en_Calc_Pcio_Costo=0 then
                    begin
                      oLista_Comprobantes_Ingresos_Sin_Subir.Add(oIngreso_Mercaderia)
                    end
                    else
                    begin
                      // Guardar en el log
                      try
                        sRespuesta:= oGestor_Lanzador_Zafiro_Local._Insertar_BI_comp_ing_stock_sin_procesar(pParametro_Zafiro_BI.Vs_Id_Cliente, oIngreso_Mercaderia.Id_empresa, oIngreso_Mercaderia.Id_sucursal, oIngreso_Mercaderia.Id_comprobante, oIngreso_Mercaderia.Fecha_comprobante, 'Comprobante con '+IntToStr(oIngreso_Mercaderia._Cantidad_Excluidos_en_Calc_Pcio_Costo)+' lineas excluidas');
                        if Length(sRespuesta)>0 then
                        begin
                          _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - '+sRespuesta);
                        end;
                      except

                      end;

                      sRespuesta:= oGestor_Lanzador_Zafiro_Local._Marcar_Como_Registrado_Comprobantes_Ingresos_En_BI(oIngreso_Mercaderia.Id_empresa, oIngreso_Mercaderia.Id_comprobante, 9);
                      if Length(sRespuesta)>0 then
                      begin
                        _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - '+sRespuesta);
                      end;
                      FreeAndNil(oIngreso_Mercaderia);
                    end;
                  end;
                end
                else
                begin
                  // Controla la cantidad de Lineas del comprobante
                  dsLin_Comprobante_Ingreso:= oGestor_Lanzador_Zafiro_Local._Buscar_Lin_Comprobantes_Ingreso_Zafiro_BI(pParametro_Zafiro_BI.Vs_Id_Cliente, FieldByName('id_empresa').AsInteger, FieldByName('id_comprobante').AsString);
                  if Assigned(dsLin_Comprobante_Ingreso)then
                  begin
                    iCant_Lineas := oGestor_Lanzador_Zafiro_Local._Buscar_Lin_Comprobantes_Ingreso_Zafiro(FieldByName('id_empresa').AsInteger, FieldByName('id_comprobante').AsString);
                    if dsLin_Comprobante_Ingreso.RecordCount <> iCant_Lineas then
                    begin
                      try
                        oIngreso_Mercaderia := oGestor_Comprobante_Stock._Buscar_Ingreso_Mercaderia(FieldByName('id_empresa').AsInteger, FieldByName('id_comprobante').AsString);
                      except
                        oIngreso_Mercaderia := Nil;
                      end;
                      if oIngreso_Mercaderia <> nil then
                      begin
                        // Marco para identificar que es por diferencia de linea.
                        oIngreso_Mercaderia.Calc_Marca := True;
                        //
                        if Assigned(oIngreso_Mercaderia.Lineas) then
                        begin
                          for iItem := 0 to oIngreso_Mercaderia.Lineas.Count-1 do
                          begin
                            if not (dsLin_Comprobante_Ingreso.Locate('item', TLin_Comprobante_Mov_Stock(oIngreso_Mercaderia.Lineas[iItem]).Item,[])) then
                            begin
                              // Utilizo Calc_No_Aprobado para identificar el item que no se encuentra
                              TLin_Comprobante_Mov_Stock(oIngreso_Mercaderia.Lineas[iItem]).Calc_Marca := True;
                              //
                            end;
                          end;
                        end;

                        if oIngreso_Mercaderia._Cantidad_Excluidos_en_Calc_Pcio_Costo=0 then
                          oLista_Comprobantes_Ingresos_Sin_Subir.Add(oIngreso_Mercaderia)
                        else
                        begin
                          // Guardar en el log
                          try
                            sRespuesta:= oGestor_Lanzador_Zafiro_Local._Insertar_BI_comp_ing_stock_sin_procesar(pParametro_Zafiro_BI.Vs_Id_Cliente, oIngreso_Mercaderia.Id_empresa, oIngreso_Mercaderia.Id_sucursal, oIngreso_Mercaderia.Id_comprobante, oIngreso_Mercaderia.Fecha_comprobante, 'Comprobante con '+IntToStr(oIngreso_Mercaderia._Cantidad_Excluidos_en_Calc_Pcio_Costo)+' lineas excluidas');
                            if Length(sRespuesta)>0 then
                            begin
                              _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - '+sRespuesta);
                            end;
                          except

                          end;

                          sRespuesta:= oGestor_Lanzador_Zafiro_Local._Marcar_Como_Registrado_Comprobantes_Ingresos_En_BI(oIngreso_Mercaderia.Id_empresa, oIngreso_Mercaderia.Id_comprobante, 9);
                          if Length(sRespuesta)>0 then
                          begin
                            _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - '+sRespuesta);
                          end;;

                          FreeAndNil(oIngreso_Mercaderia);
                        end;

                      end;
                    end
                    else
                    begin
                      // El Comprobante tiene la cantidad de Lineas correcta
                      sRespuesta:= oGestor_Lanzador_Zafiro_Local._Marcar_Como_Registrado_Comprobantes_Ingresos_En_BI(FieldByName('id_empresa').AsInteger, FieldByName('id_comprobante').AsString, 1);
                      if Length(sRespuesta)>0 then
                      begin
                        _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - '+sRespuesta);
                      end;
                    end;
                  end;
                end;
              end;

              Next;
            end;
          end;
          //Close;
        end;

        if oLista_Comprobantes_Ingresos_Sin_Subir.Count > 0 then
        begin
          _Grabar_LogErrores_Comunicacion(-5, '      Fecha: '+ FormatDateTime('dd/mm/yyyy', dtFecha_Comp_Ingreso)+' Lote:'+IntToStr(iLote)+' - Se Prepararon '+ IntToStr(oLista_Comprobantes_Ingresos_Sin_Subir.Count)+' Comprobantes de Ingresos en BI - '+FormatDateTime('hh:mm:ss', Now)+' Duración: '+FormatDateTime('hh:mm:ss', (Now-dtFechaHora_Ini)));
          if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
          begin
            _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Reconexión con la BD Zafiro BI...');
            if Assigned(oGestor_Lanzador_Zafiro_Local) then
              FreeAndNil(oGestor_Lanzador_Zafiro_Local);
            oGestor_Lanzador_Zafiro_Local:= TGestor_Lanzador_Zafiro.Create(pParametro_Zafiro_BI.Cadena_Conexion_BD);
            if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
              _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - No se pudo establecer conexión con la BD Zafiro BI. (8)');
          end;

          if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI then
          begin
            bError:= False;
            sResultado := oGestor_Lanzador_Zafiro_Local._Subir_Comprobantes_Ingresos_Zafiro_BI(pParametro_Zafiro_BI.Vs_Id_Cliente, oLista_Comprobantes_Ingresos_Sin_Subir, oGestor_Proveedores);

            if Length(Trim(sResultado))>0 then
            begin
              _Grabar_LogErrores_Comunicacion(-5, '   Fecha: '+ FormatDateTime('dd/mm/yyyy', dtFecha_Comp_Ingreso)+' - '+ sResultado);
              if Length(Result)>0 then
                Result:= Result + Chr(13);
              Result:= Result + sResultado;
              bError:= True;
            end
            else
            begin
              if bError=False then
              begin
                dtFechaHora_Fin:= Now;
                _Grabar_LogErrores_Comunicacion(-5, '      Fecha: '+ FormatDateTime('dd/mm/yyyy', dtFecha_Comp_Ingreso)+' Lote:'+IntToStr(iLote)+' - Se Registraron '+ IntToStr(oLista_Comprobantes_Ingresos_Sin_Subir.Count)+' Comprobantes de Ingreso en BI - '+FormatDateTime('hh:mm:ss', dtFechaHora_Fin)+' Duración Total: '+FormatDateTime('hh:mm:ss', (dtFechaHora_Fin-dtFechaHora_Ini)));
              end;
            end;
            iLote:= iLote + 1;

          end;
          {for iItem := 0 to oLista_Comprobantes_Ingresos_Sin_Subir.Count-1 do
          begin
            if Assigned(TComprobante_Ventas(oLista_Comprobantes_Ingresos_Sin_Subir[iItem])) then
            begin
              if TComprobante_Ventas(oLista_Comprobantes_Ingresos_Sin_Subir[iItem]).Tipo_Comprobante=NC__NotaDeCredito then
              begin
                oNota_de_Credito := TNota_de_Credito(oLista_Comprobantes_Ingresos_Sin_Subir[iItem]);
                FreeAndNil(oNota_de_Credito);
              end
              else
              begin
                oFactura_ND_Presupuesto := TFactura_ND_Presupuesto(oLista_Comprobantes_Ingresos_Sin_Subir[iItem]);
                FreeAndNil(oFactura_ND_Presupuesto);
              end;
            end;
          end;}


        end;

        // Destruye la lista de Comprobantes
        //oLista_Comprobantes_Ingresos_Sin_Subir.Free;
        FreeAndNil(oLista_Comprobantes_Ingresos_Sin_Subir);
        //
        //FreeandNil(dsComprobantes_Ingresos);
        //
        //if dsComprobantes_Ingresos.Active then
        //  dsComprobantes_Ingresos.Close;

        //LiberarMemoria;

        {
        // Para reiniciar cada 5 lotes
        if iLote mod 5 = 0 then
        begin
          if Assigned(oGestor_Lanzador_Zafiro_Local) then
            FreeAndNil(oGestor_Lanzador_Zafiro_Local);

          if Assigned(oGestor_Comprobante_Ventas) then
            FreeAndNil(oGestor_Comprobante_Ventas);

          if Assigned(oGestor_Vendedores) then
            FreeAndNil(oGestor_Vendedores);

          if Assigned(oGestor_Clientes) then
            FreeAndNil(oGestor_Clientes);
          Result:= 'Reiniciar';
          Exit;
        end;
        //}

      end
      else
      begin
        bHayComprobantes:= False;
        iLote:= 1;
      end;

    end;

    if Length(Result)=0 then
    begin

      //if not Assigned(oGestor_Lanzador_Zafiro_Local) then
      //  oGestor_Lanzador_Zafiro_Local:= TGestor_Lanzador_Zafiro.Create;

      sResultado := oGestor_Lanzador_Zafiro_Local._Actualizar_Fecha_Ult_Act_Comprobantes_Ingresos_BI(dtFecha_Comp_Ingreso);
      if Length(Trim(sResultado))>0 then
      begin
        _Grabar_LogErrores_Comunicacion(-5, '   Fecha: '+ FormatDateTime('dd/mm/yyyy', dtFecha_Comp_Ingreso)+' - '+ sResultado);
        if Length(Result)>0 then
          Result:= Result + Chr(13);
        Result:= Result + sResultado;
      end;
    end;

    dtFecha_Comp_Ingreso := dtFecha_Comp_Ingreso + 1;
  end; // while dtFecha_Comp_Ingreso<=pFecha_Hasta do


  if Assigned(oGestor_Proveedores) then
    FreeAndNil(oGestor_Proveedores);

  ///////////////////////////////////////////
  // Fin Comprobantes de Ingresos          //
  ///////////////////////////////////////////


  ///////////////////////////////////////////
  // Stock Por Mes                         //
  ///////////////////////////////////////////
  // Meses Anteriores
  for iCantidad_Meses_Hacia_Atras := 1 to iCantidad_Meses_Hacia_Atras do
  begin
    dtFecha_Stock_Por_Mes := DateOf(TDate(EndOfTheMonth(IncMonth(pFecha_Hasta, iCantidad_Meses_Hacia_Atras * (-1)))));

    if oGestor_Lanzador_Zafiro_Local._Existe_Stock_Por_Mes_Log_BI(pParametro_Zafiro_BI.Vs_Id_Cliente, dtFecha_Stock_Por_Mes, 'ME')=False then
    begin
      // Subir Stock Por Mes

      try
        dtFecha_Vtas_Desde := dtFecha_Stock_Por_Mes - iCantidad_Dias_Vtas_Ant + 1;
        dtFecha_Vtas_Hasta := dtFecha_Stock_Por_Mes;
        dsStock_Por_Mes := oGestor_Lanzador_Zafiro_Local._Buscar_Stock_Por_Mes_Zafiro_BI(dtFecha_Stock_Por_Mes, dtFecha_Vtas_Desde, dtFecha_Vtas_Hasta, iCantidad_Dias_Vtas_Ant);
      except
      end;

      if Assigned(dsStock_Por_Mes) then
      begin
        with dsStock_Por_Mes do
        begin
          if RecordCount > 0 then
          begin
            _Grabar_LogErrores_Comunicacion(-5, '   Registrando '+ IntToStr(RecordCount)+' Registros de Stock por Mes del día '+  FormatDateTime('yyyy/mm/dd',  dtFecha_Stock_Por_Mes) +' en Zafiro BI - '+FormatDateTime('hh:mm:ss', Now));
            if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
            begin
              _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Reconexión con la BD Zafiro BI...');
              if Assigned(oGestor_Lanzador_Zafiro_Local) then
                FreeAndNil(oGestor_Lanzador_Zafiro_Local);
              oGestor_Lanzador_Zafiro_Local:= TGestor_Lanzador_Zafiro.Create(pParametro_Zafiro_BI.Cadena_Conexion_BD);
              if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
                _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - No se pudo establecer conexión con la BD Zafiro BI. (5)');
            end;

            if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI then
            begin
              sResultado := oGestor_Lanzador_Zafiro_Local._Subir_Stock_Por_Mes_Zafiro_BI_Zafiro_BI(pParametro_Zafiro_BI.Vs_Id_Cliente, dsStock_Por_Mes, dtFecha_Stock_Por_Mes, 'ME', iCantidad_Dias_Vtas_Ant, 100);
              if Length(Trim(sResultado))>0 then
              begin
                _Grabar_LogErrores_Comunicacion(-5, '   '+sResultado);
              end
              else
              begin
                _Grabar_LogErrores_Comunicacion(-5, '      Se Registraron '+ IntToStr(RecordCount)+' Registros de Stock por Mes del día '+  FormatDateTime('yyyy/mm/dd',  dtFecha_Stock_Por_Mes) +' en Zafiro BI - '+FormatDateTime('hh:mm:ss', Now));
              end;
            end;
          end;
        end;
      end;

      //
    end;

  end;
  ///////////////////////////////////////////
  // Fin Stock Por Mes                     //
  ///////////////////////////////////////////

  LiberarMemoria;

  ///////////////////////////////////////////
  // Stock Del Ultimo dia                  //
  ///////////////////////////////////////////

  if oGestor_Lanzador_Zafiro_Local._Existe_Stock_Por_Mes_Log_BI(pParametro_Zafiro_BI.Vs_Id_Cliente, pFecha_Hasta, 'UD')=False then
  begin
    sResultado := oGestor_Lanzador_Zafiro_Local._Eliminar_Stock_Por_Mes_Log_BI_Ultimo_Dia(pParametro_Zafiro_BI.Vs_Id_Cliente);
    if Length(Trim(sResultado))>0 then
    begin
      _Grabar_LogErrores_Comunicacion(-5, '   '+sResultado);
    end
    else
    begin
      _Grabar_LogErrores_Comunicacion(-5, '      Se elimino registros de stock del último día en Zafiro BI - '+FormatDateTime('hh:mm:ss', Now));
      // Subir Stock Ultimo dia

      // Le asigno a la fecha pasada por parametros que es la del dia anterior
      dtFecha_Stock_Por_Mes := pFecha_Hasta;

      try
        dtFecha_Vtas_Desde := dtFecha_Stock_Por_Mes - iCantidad_Dias_Vtas_Ant + 1;
        dtFecha_Vtas_Hasta := dtFecha_Stock_Por_Mes;
        dsStock_Por_Mes := oGestor_Lanzador_Zafiro_Local._Buscar_Stock_Por_Mes_Zafiro_BI(dtFecha_Stock_Por_Mes, dtFecha_Vtas_Desde, dtFecha_Vtas_Hasta, iCantidad_Dias_Vtas_Ant);
      except
      end;

      if Assigned(dsStock_Por_Mes) then
      begin
        with dsStock_Por_Mes do
        begin
          if RecordCount > 0 then
          begin
            _Grabar_LogErrores_Comunicacion(-5, '   Registrando '+ IntToStr(RecordCount)+' Registros de Stock por Mes del día '+  FormatDateTime('yyyy/mm/dd',  dtFecha_Stock_Por_Mes) +' en Zafiro BI - '+FormatDateTime('hh:mm:ss', Now));
            if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
            begin
              _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Reconexión con la BD Zafiro BI...');
              if Assigned(oGestor_Lanzador_Zafiro_Local) then
                FreeAndNil(oGestor_Lanzador_Zafiro_Local);
              oGestor_Lanzador_Zafiro_Local:= TGestor_Lanzador_Zafiro.Create(pParametro_Zafiro_BI.Cadena_Conexion_BD);
              if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI=False then
                _Grabar_LogErrores_Comunicacion(-5, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - No se pudo establecer conexión con la BD Zafiro BI. (5)');
            end;

            if oGestor_Lanzador_Zafiro_Local._Hay_Conexion_Zafiro_BI then
            begin
              sResultado := oGestor_Lanzador_Zafiro_Local._Subir_Stock_Por_Mes_Zafiro_BI_Zafiro_BI(pParametro_Zafiro_BI.Vs_Id_Cliente, dsStock_Por_Mes, dtFecha_Stock_Por_Mes, 'UD', iCantidad_Dias_Vtas_Ant, 100);
              if Length(Trim(sResultado))>0 then
              begin
                _Grabar_LogErrores_Comunicacion(-5, '   '+sResultado);
              end
              else
              begin
                _Grabar_LogErrores_Comunicacion(-5, '      Se Registraron '+ IntToStr(RecordCount)+' Registros de Stock por Mes del día '+  FormatDateTime('yyyy/mm/dd',  dtFecha_Stock_Por_Mes) +' en Zafiro BI - '+FormatDateTime('hh:mm:ss', Now));
              end;
            end;
          end;
        end;
      end;

      //
    end;
  end;
  ///////////////////////////////////////////
  // Fin Stock Del Ultimo dia              //
  ///////////////////////////////////////////

  LiberarMemoria;

  if Assigned(oGestor_Lanzador_Zafiro_Local) then
    FreeAndNil(oGestor_Lanzador_Zafiro_Local);

end;

procedure TVSComunicacionParaReplicacion01._Agregar_CadenaSQL(pSucursal: TSucursal; pCadena: String; pPosicion_caracter_final: integer; pTabla_Exportar: String = '');
Var
  sNombreTabla: String;
begin
  sNombreTabla:= _Extraer_NombreTabla(pCadena, pPosicion_caracter_final);
  {
  //guardo el nombre de la tabla
  while pPosicion_caracter_final<=(Length(pCadena)) do //mientras que no se llegue a la posición final de la cadena
  begin

    if ((pCadena[pPosicion_caracter_final]=' ') or (pCadena[pPosicion_caracter_final]=#13) or (pCadena[pPosicion_caracter_final]=#10) or (pCadena[pPosicion_caracter_final]='(')) then
    begin
      //Si ya encontre el nombre de la tabla salgo del bucle
      if Length(sNombreTabla) > 0 then
      begin
        sNombreTabla:= StringReplace(sNombreTabla,'`','', [rfReplaceAll]);
        Break;
      end;
    end
    else
    begin
      sNombreTabla:= sNombreTabla + pCadena[pPosicion_caracter_final];
    end;
    pPosicion_caracter_final := pPosicion_caracter_final + 1;
  end; //end while pPosicion_caracter_final<=(Length(pCadena)) do //mientras que no se llegue a la posición final de la cadena
  }

  if Length(sNombreTabla) > 0 then
  begin
    //-------------------------
    // ojo pv
    // Version para sucursales
    // Comentar cuando no lo sea
    // Para Sucursal: Sin comentar
    // Para Central: Comentar
    //-------------------------
//    if sNombreTabla='log_comunicacion_externo' then
//    begin
//      slSentencias_comunicacion.Clear;
//      exit;
//    end;
    //-------------------------

    //Si la sucursal no tiene como exceptuada la tabla que guarde en sNombreTabla recién guardo la sentencia
    if _Tabla_Exceptuada(sNombreTabla, pSucursal)=False then
    begin
      if Length(pTabla_Exportar)=0 then
        slSentencias_comunicacion.Add(pCadena)
      else
        if LowerCase(pTabla_Exportar)=LowerCase(sNombreTabla) then
           _Grabar_LogErrores_Comunicacion(-99, pCadena+';');
    end;

    //if Assigned(pSucursal.Tablas_exceptuadas_sucursal) then
    //begin
    //  if not (pSucursal.Tablas_exceptuadas_sucursal.IndexOf(sNombreTabla)> -1) then
    //  begin
    //    slSentencias_comunicacion.Add(pCadena);
    //  end;
    //end
    //else
    //begin
    //  slSentencias_comunicacion.Add(pCadena );
    //end;

  end;//end if Length(sNombreTabla) > 0 then
end;

function TVSComunicacionParaReplicacion01._Archivo_Cargado_FTP(pArchivo: string;
  pLista: TStringList): Boolean;
var
  iItem: Integer;
begin
  Result := False;
  if Assigned(pLista) then
  begin
    if pLista.Count > 0 then
    begin
      for iItem := 0 to pLista.Count - 1 do
      begin
        if pLista[iItem] = pArchivo then
        begin
          Result := True;
          Break;
        end;
      end;
    end;
  end;
end;

procedure TVSComunicacionParaReplicacion01._Grabar_LogErrores_Comunicacion(pId_ProcesoActualizacion: Smallint; pCadenaError: AnsiString; pId_Sucursal: Integer=0);
const
  Ruta = 'c:\vs-comunicacion\';
var
  sFileName: String;
  F: TextFile;
begin
  //****************Guardo en un archivo log el error de transmisión ***********************//
  if pId_ProcesoActualizacion>=0 then
    sFileName:= 'Log_Error_' + FormatDateTime('yyyy-MM-dd', DateOf(Now)) + '-'+IntToStr(pId_ProcesoActualizacion)+'.txt'
  else if pId_ProcesoActualizacion=-1 then
    sFileName:= 'Log_Etiquetas' + FormatDateTime('yyyy-MM-dd', DateOf(Now)) + '.txt'
  else if pId_ProcesoActualizacion=-2 then
    sFileName:= 'Log_IQVia' + FormatDateTime('yyyy-MM-dd', DateOf(Now)) + '.txt'
  else if pId_ProcesoActualizacion=-3 then
    sFileName:= 'Log_CloseUp' + FormatDateTime('yyyy-MM-dd', DateOf(Now)) + '.txt'
  else if pId_ProcesoActualizacion=-4 then
    sFileName:= 'Log_ActuPreciosProgramada' + FormatDateTime('yyyy-MM-dd', DateOf(Now)) + '.txt'
  else if pId_ProcesoActualizacion=-5 then
    sFileName:= 'Log_ZafiroBI' + FormatDateTime('yyyy-MM-dd', DateOf(Now)) + '.txt'
  else if pId_ProcesoActualizacion=-6 then
    sFileName:= 'Log_SubidaCompFTP' + FormatDateTime('yyyy-MM-dd', DateOf(Now)) + '.txt'
  else if pId_ProcesoActualizacion=-7 then
    sFileName:= 'Log_Actualiz_MinMax_Stock' + FormatDateTime('yyyy-MM-dd', DateOf(Now)) + '.txt'
  else if pId_ProcesoActualizacion=-8 then
    sFileName:= 'Log_Transmision_de_Sentencias_Por_Comprobante' + FormatDateTime('yyyy-MM-dd', DateOf(Now))+ '-'+IntToStr(pId_Sucursal) + '.txt'
  else if pId_ProcesoActualizacion=-9 then
    sFileName:= 'Log_ActuPreciosCompraPP' + FormatDateTime('yyyy-MM-dd', DateOf(Now)) + '.txt'
  else if pId_ProcesoActualizacion=-10 then
    sFileName:= 'Log_Exportacion_DataView' + FormatDateTime('yyyy-MM-dd', DateOf(Now)) + '.txt'
  else if pId_ProcesoActualizacion=-99 then
    sFileName:= 'Sentencias_Tabla' + FormatDateTime('yyyy-MM-dd', DateOf(Now)) + '.sql';

  AssignFile(f,Ruta+sFileName);
  if FileExists(Ruta+sFileName) then
    Append(f)
  else
    Rewrite(f);
  writeln(f,pCadenaError );
  CloseFile(f);
  //****************Guardo en un archivo log el error de transmisión ***********************//
end;




function TVSComunicacionParaReplicacion01._Existe_log_Posterior(pId_archivo_log_procesado: String): Boolean;
var
  iNro_Extension: Integer;
begin
  iNro_Extension:= StrToInt(RightStr(pId_archivo_log_procesado, oParametro.Comu_largo_extension_file_binlog)) + 1;
  pId_archivo_log_procesado:= oParametro.Comu_nombre_file_binlog + '.'+ RightStr('00000000000000000'+IntToStr(iNro_Extension), oParametro.Comu_largo_extension_file_binlog);

  with qrySelect_Binlog_events do
  begin
    sql.Clear;
    sql.Add('SHOW BINLOG EVENTS IN '+ QuotedStr(pId_archivo_log_procesado)+' LIMIT 10');
    //sql.Add('SELECT * FROM bin_log_events_tabla WHERE bin_log_events_tabla.Log_name= '+ QuotedStr(pId_archivo_log_procesado) +' LIMIT 10');
    try
      Open;
      if RecordCount > 0 then
        Result:= True
      else
        Result:= False;
    except
      Result:= False;
    end;
    close;
  end;//end with qrySelect_Binlog_events do
end;

function TVSComunicacionParaReplicacion01._Carga_Sucursales(pId_ProcesoActualizacion: Smallint; pErrorComunicacion: Boolean):TObjectList;
var
  oSucursales_local: TObjectList;
  oSucursal: TSucursal;
  sCadena_error: String;
  oTablas_exceptuadas_sucursal_local: TStringList;
begin
  try
    with qrySelectAll_sucursales_comunicacion do
    begin
      Close;
      Open;
      First;
      oSucursales_local := TObjectList.Create(True);
      while not Eof do
      begin
        if ((FieldByName('Id_Sucursal').AsInteger <> oParametro.Comu_Id_Sucursal) and (FieldByName('Exceptuado').AsInteger<> 1) and (Length(FieldByName('Cadena_conexion').AsString)<> 0) ) then
        begin

          // Eliminar esta pregunta y el parametro pId_ProcesoActualizacion cunado se unifique el proceso en un solo servicio
          if FieldByName('Id_ProcesoActualizacion').AsInteger=pId_ProcesoActualizacion then
          begin

            oSucursal := Nil;
            oSucursal := TSucursal.Create(Fields);
            oTablas_exceptuadas_sucursal_local:= TStringList.Create;

            //***********Cargo la lista de tablas exceptuadas de la sucursal**************//
            qrySelect_tablas_exceptuadas_por_sucursal.Close;
            qrySelect_tablas_exceptuadas_por_sucursal.Parameters.ParamByName('pId_sucursal').Value:= oSucursal.Id_sucursal;
            qrySelect_tablas_exceptuadas_por_sucursal.Open;
            if qrySelect_tablas_exceptuadas_por_sucursal.RecordCount> 0 then
            begin
              qrySelect_tablas_exceptuadas_por_sucursal.First;
              while not qrySelect_tablas_exceptuadas_por_sucursal.Eof do
              begin
                //oSucursal.Tablas_exceptuadas_sucursal.Add(qrySelect_tablas_exceptuadas_por_sucursal.FieldByName('tabla').AsString);
                oTablas_exceptuadas_sucursal_local.Add(qrySelect_tablas_exceptuadas_por_sucursal.FieldByName('tabla').AsString);
                qrySelect_tablas_exceptuadas_por_sucursal.Next;
              end;//end while not Eof do
            end;//end if qrySelect_tablas_exceptuadas_por_sucursal.RecordCount> 0 then
            qrySelect_tablas_exceptuadas_por_sucursal.Close;
            //************Cargo la lista de tablas exceptuadas de la sucursal**************//
            oSucursal.Tablas_exceptuadas_sucursal := oTablas_exceptuadas_sucursal_local;
            oSucursales_local.add(oSucursal);
          end;
        end;
        Next;
      end; //end while not Eof do
    end; //end with qrySelectAll_sucursales do
    Result:= oSucursales_local;
  except
    on E: Exception do
    begin
      //Utilizo esta variable bErrorComunicacion para grabar solo una vez el error de comunicacion si el problema persiste
      if pErrorComunicacion <> true then
      begin
        sCadena_error:= DateTimeToStr(Now)+' -  Error: '+ (E.Message);
        _Grabar_LogErrores_Comunicacion(0, sCadena_error);
      end;
      Result:= Nil;
    end;
  end;
end;

procedure TVSComunicacionParaReplicacion01._Conectar_Servidor_FTP(pHost, pUser,
  pPass: String; pPort, pTimeout: Integer; pModoPasivo : Boolean = False);
begin
  if IdFTP1.Connected then
  begin
    IdFTP1.Disconnect;
  end;

  IdFTP1.Host            := pHost;
  IdFTP1.Port            := pPort;    // Generalmente = 21
  IdFTP1.TransferTimeout := pTimeout; // Tiempo de Respuesta expresado en milisegundos
  IdFTP1.Username        := pUser;
  IdFTP1.Password        := pPass;

  Try
    _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Intentando conectar a ' + pHost );
     if pModoPasivo then
     begin
      IdFTP1.Passive := True;
      IdFTP1.Connect;
      _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' -  Conectado a ' + pHost  + ' Con Modo Pasivo Activado');
     end
     else
     begin
      IdFTP1.Connect;
      _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' -  Conectado a ' + pHost );
     end;

  Except
     //Esta clase de exepción cubre los casos del lado del servidor cierran la conexión
    on EError: EIdConnClosedGracefully do
      _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Falló intentando conectar a ' + pHost  + ' Error: ' + 'Conexión cerrada por el Servidor');

    // Esta clase de exepción cubre los casos de errores de socket de bajo nivel
    on EError: EIdSocketError do
      _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Falló intentando conectar a ' + pHost  + ' Error de Socket: ' + sLineBreak +
        'Error code: ' + IntToStr(EError.LastError) + sLineBreak +
        'Error message' + EError.Message);

    // Esta clase de exepción cubre los casos de errores que lanza la Librería Indy
    on EError: EIdException do
      _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Falló intentando conectar a ' + pHost  + ' Error Librería Indy: ' + sLineBreak +
        'Exception class: ' + EError.ClassName + sLineBreak +
        'Error message: ' + EError.Message);

    // Esta clase de exepción cubre los casos de errores de otros componentes no listados en los casos anteriores.
      on EError: Exception do
        _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Falló intentando conectar a ' + pHost  + ' Error: ' +EError.Message);
  end;


  //Si Falló al conectar, hago un segundo intento utilizando Passivemode

  if not(IdFTP1.Connected) then
  Try
    _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Intentando conectar a ' + pHost  + ' Con Modo Pasivo Activado');

    IdFTP1.Connect;
    _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' -  Conectado a ' + pHost  + ' Con Modo Pasivo Activado');

  Except
    on EError: Exception do
      _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Falló intentando conectar a ' + pHost  + ' Error: ' +EError.Message);

     //Esta clase de exepción cubre los casos del lado del servidor cierran la conexión
    on EError: EIdConnClosedGracefully do
      _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Falló intentando conectar a ' + pHost  + ' Error: ' + 'Conexión cerrada por el Servidor');

    // Esta clase de exepción cubre los casos de errores de socket de bajo nivel
    on EError: EIdSocketError do
      _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Falló intentando conectar a ' + pHost  + ' Error de Socket: ' + sLineBreak +
        'Error code: ' + IntToStr(EError.LastError) + sLineBreak +
        'Error message' + EError.Message);

    // Esta clase de exepción cubre los casos de errores que lanza la Librería Indy
    on EError: EIdException do
      _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Falló intentando conectar a ' + pHost  + ' Error Librería Indy: ' + sLineBreak +
        'Exception class: ' + EError.ClassName + sLineBreak +
        'Error message: ' + EError.Message);

    // Esta clase de exepción cubre los casos de errores de otros componentes no listados en los casos anteriores.
      on EError: Exception do
        _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Falló intentando conectar a ' + pHost  + ' Error: ' +EError.Message);

  end;
end;

Function TVSComunicacionParaReplicacion01._Confecciona_Linea_Sentencias_Faltantes(pTabla: String; pSentencia: String; var pListaSentencias: TStringList; pSucursal: TSucursal; pTipo_Sentencia: String ='INSERT'): Boolean;
var
  bPrimera: Boolean;
  Formato : TFormatSettings;
  sSentencia_local: String;
  sSentencia_local_Primera_Parte: String;
  iItem_Lote: SmallInt;
  bVaColumna: Boolean;
begin
  Result:= False;
  if _Tabla_Exceptuada(pTabla, pSucursal)=True then
    Exit;

  Formato.DecimalSeparator := '.';
  Formato.ThousandSeparator := #0;

  if pTipo_Sentencia='INSERT' then
  begin

    qryDescribe_tabla.Close;
    qryDescribe_tabla.SQL.Clear;
    qryDescribe_tabla.SQL.Add('DESCRIBE '+pTabla);
    qryDescribe_tabla.Open;
    qryDescribe_tabla.First;

    bPrimera:= True;
    if qryDescribe_tabla.RecordCount>0 then
    begin
      sSentencia_local_Primera_Parte:= 'INSERT INTO '+pTabla+' (';
      while Not qryDescribe_tabla.Eof do
      begin

        bVaColumna:= True;
        if oParametro.Estructura_Coseguro_orden_cupones_entidad_convenio=False then
        begin
          if ((qryDescribe_tabla.FieldByName('Field').AsString='imp_uni_coseguro1_civa') or
            (qryDescribe_tabla.FieldByName('Field').AsString='imp_uni_coseguro2_civa') or
            (qryDescribe_tabla.FieldByName('Field').AsString='imp_uni_coseguro3_civa')) then
          begin
            bVaColumna:= False;
          end;
        end;

        if bVaColumna then
        begin
          if bPrimera=False then
            sSentencia_local_Primera_Parte:=sSentencia_local_Primera_Parte + ','
          else
            bPrimera:= False;
          sSentencia_local_Primera_Parte:= sSentencia_local_Primera_Parte + qryDescribe_tabla.FieldByName('Field').AsString;
        end;
        qryDescribe_tabla.Next;
      end;
      sSentencia_local_Primera_Parte:=sSentencia_local_Primera_Parte+') VALUES ';
    end
    else
    begin
      Result:= False;
      Exit;
    end;

    //pListaSentencias.Add(sSentencia_local_Primera_Parte);

    sSentencia_local:=sSentencia_local_Primera_Parte;

    iItem_Lote:= 0;
    qrySelectManual.SQL.Clear;
    qrySelectManual.SQL.Add(pSentencia);
    qrySelectManual.Open;
    qrySelectManual.First;
    if qrySelectManual.RecordCount>0 then
    begin
      while Not qrySelectManual.Eof do
      begin

        qryDescribe_tabla.First;
        if qryDescribe_tabla.RecordCount>0 then
        begin
          bPrimera:= True;
          while Not qryDescribe_tabla.Eof do
          begin

            bVaColumna:= True;
            if oParametro.Estructura_Coseguro_orden_cupones_entidad_convenio=False then
            begin
              if ((qryDescribe_tabla.FieldByName('Field').AsString='imp_uni_coseguro1_civa') or
                (qryDescribe_tabla.FieldByName('Field').AsString='imp_uni_coseguro2_civa') or
                (qryDescribe_tabla.FieldByName('Field').AsString='imp_uni_coseguro3_civa')) then
              begin
                bVaColumna:= False;
              end;
            end;

            if bPrimera=False then
            begin
              if bVaColumna then
              begin
                sSentencia_local:=sSentencia_local + ',';
              end;
            end
            else
            begin
              bPrimera:= False;
              sSentencia_local:=sSentencia_local + '('
            end;

            if bVaColumna then
            begin
              if qrySelectManual.FieldByName(qryDescribe_tabla.FieldByName('Field').AsString).Value = null then
              begin
                sSentencia_local:= sSentencia_local + 'NULL';
              end
              // Si es id_pedido de lin_comprobantes_de_ventas
              else if ((pTabla= 'lin_comprobantes_de_ventas') and (qryDescribe_tabla.FieldByName('Field').AsString='id_pedido')) then
              begin
                sSentencia_local:= sSentencia_local + 'NULL';
              end
              //
              else
              begin
                if LeftStr(qryDescribe_tabla.FieldByName('Type').AsString,4) ='int(' then // Entero
                begin
                  sSentencia_local:= sSentencia_local + qrySelectManual.FieldByName(qryDescribe_tabla.FieldByName('Field').AsString).AsString;
                end
                else if LeftStr(qryDescribe_tabla.FieldByName('Type').AsString,9) ='smallint(' then // smallint
                begin
                  sSentencia_local:= sSentencia_local + qrySelectManual.FieldByName(qryDescribe_tabla.FieldByName('Field').AsString).AsString;
                end
                else if LeftStr(qryDescribe_tabla.FieldByName('Type').AsString,7) ='tinyint' then // tinyint
                begin
                  sSentencia_local:= sSentencia_local + qrySelectManual.FieldByName(qryDescribe_tabla.FieldByName('Field').AsString).AsString;
                end
                else if LeftStr(qryDescribe_tabla.FieldByName('Type').AsString,6) ='bigint' then // bigint
                begin
                  sSentencia_local:= sSentencia_local + qrySelectManual.FieldByName(qryDescribe_tabla.FieldByName('Field').AsString).AsString;
                end
                else if LeftStr(qryDescribe_tabla.FieldByName('Type').AsString,6) ='double' then // double
                begin
                  sSentencia_local:= sSentencia_local + FloatToStr(qrySelectManual.FieldByName(qryDescribe_tabla.FieldByName('Field').AsString).AsFloat, Formato);
                end
                else if LeftStr(qryDescribe_tabla.FieldByName('Type').AsString,8) ='varchar(' then // Varchar
                begin
                  sSentencia_local:= sSentencia_local + QuotedStr(qrySelectManual.FieldByName(qryDescribe_tabla.FieldByName('Field').AsString).AsString);
                end
                else if LeftStr(qryDescribe_tabla.FieldByName('Type').AsString,5) ='char(' then // Varchar
                begin
                  sSentencia_local:= sSentencia_local + QuotedStr(qrySelectManual.FieldByName(qryDescribe_tabla.FieldByName('Field').AsString).AsString);
                end
                else if LeftStr(qryDescribe_tabla.FieldByName('Type').AsString,4) ='text' then // text
                begin
                  sSentencia_local:= sSentencia_local + QuotedStr(qrySelectManual.FieldByName(qryDescribe_tabla.FieldByName('Field').AsString).AsString);
                end
                else if LeftStr(qryDescribe_tabla.FieldByName('Type').AsString,9) ='timestamp' then // timestamp
                begin
                  sSentencia_local:= sSentencia_local + QuotedStr(FormatDateTime('yyyy-mm-dd HH:mm:ss', qrySelectManual.FieldByName(qryDescribe_tabla.FieldByName('Field').AsString).AsDateTime));
                end
                else if LeftStr(qryDescribe_tabla.FieldByName('Type').AsString,4) ='date' then // timestamp
                begin
                  sSentencia_local:= sSentencia_local + QuotedStr(FormatDateTime('yyyy-mm-dd', qrySelectManual.FieldByName(qryDescribe_tabla.FieldByName('Field').AsString).AsDateTime));
                end
                else // Cualquier otra cosa como texto
                begin
                  sSentencia_local:= sSentencia_local + QuotedStr(qrySelectManual.FieldByName(qryDescribe_tabla.FieldByName('Field').AsString).AsString);
                end
              end;

            end;
            qryDescribe_tabla.Next;
          end;
        end;
        sSentencia_local:= sSentencia_local +')';

        iItem_Lote:= iItem_Lote + 1;
        qrySelectManual.Next;
        if qrySelectManual.Eof then
        begin
          sSentencia_local:= sSentencia_local +';';
          pListaSentencias.Add(sSentencia_local);
        end
        else
        begin
          if iItem_Lote>=20 then
          begin
            sSentencia_local:= sSentencia_local +';';
            pListaSentencias.Add(sSentencia_local);
            //pListaSentencias.Add(sSentencia_local_Primera_Parte);
            sSentencia_local:=sSentencia_local_Primera_Parte;
            iItem_Lote:=0;
          end
          else
          begin
            sSentencia_local:= sSentencia_local +','+chr(13);
          end;
        end;
      end;
    end
    else
    begin
      Result:= False;
    end;


  end
  else if pTipo_Sentencia='UPDATE' then // SENTENCIA
  begin

    qrySelectManual.SQL.Clear;
    qrySelectManual.SQL.Add(pSentencia);
    qrySelectManual.Open;
    qrySelectManual.First;
    if qrySelectManual.RecordCount>0 then
    begin
      while Not qrySelectManual.Eof do
      begin
        qrySelectManual.Fields[0].Value;
        pListaSentencias.Add(qrySelectManual.Fields[0].AsString);
        qrySelectManual.Next;
      end;
    end
    else
    begin
      Result:= False;
    end;
  end
  else if pTipo_Sentencia='SENTENCIA MANUAL' then // SENTENCIA
  begin
    pListaSentencias.Add(pSentencia);
  end;
end;

function TVSComunicacionParaReplicacion01._Confecciona_Sentencias_Faltantes(
  pTipo_Comprobante: String; pParamatro1: String; pParamatro2: String; pParamatro3: String; pParamatro4: String;
  pSucursal: TSucursal): TStringList;
var
  bResultado: Boolean;
  oNota_de_Credito: TNota_de_Credito;
begin
  Result:= Nil;
  if pTipo_Comprobante= 'pedido' then
  begin
    Result:= TStringList.Create;
      if pParamatro1<>'0' then
      begin
        _Confecciona_Linea_Sentencias_Faltantes('pedidos'   , 'SELECT pedidos.*    FROM pedidos    WHERE pedidos.id_empresa='+pParamatro1+' AND pedidos.id_pedido='+ QuotedStr(pParamatro2), Result, pSucursal);
        _Confecciona_Linea_Sentencias_Faltantes('lin_pedido', 'SELECT lin_pedido.* FROM lin_pedido WHERE lin_pedido.id_empresa='+pParamatro1+' AND lin_pedido.id_pedido='+ QuotedStr(pParamatro2), Result, pSucursal);
      end
      else
      begin
        _Confecciona_Linea_Sentencias_Faltantes('pedidos'   , 'SELECT pedidos.*    FROM pedidos    WHERE pedidos.id_pedido='+ QuotedStr(pParamatro2), Result, pSucursal);
        _Confecciona_Linea_Sentencias_Faltantes('lin_pedido', 'SELECT lin_pedido.* FROM lin_pedido WHERE lin_pedido.id_pedido='+ QuotedStr(pParamatro2), Result, pSucursal);
      end;
  end
  else if pTipo_Comprobante= 'factura' then
  begin
    Result:= TStringList.Create;
    _Confecciona_Linea_Sentencias_Faltantes('comprobantes_de_ventas'        , 'SELECT comprobantes_de_ventas.*         FROM comprobantes_de_ventas         WHERE comprobantes_de_ventas.id_empresa='+pParamatro1+' AND comprobantes_de_ventas.id_comprobante='+ QuotedStr(pParamatro2), Result, pSucursal);
    if oParametro.Estructura_Id_empresa_lin_comprobantes_de_ventas then
      _Confecciona_Linea_Sentencias_Faltantes('lin_comprobantes_de_ventas'    , 'SELECT lin_comprobantes_de_ventas.*     FROM lin_comprobantes_de_ventas     WHERE lin_comprobantes_de_ventas.id_empresa='+pParamatro1+' AND lin_comprobantes_de_ventas.id_comprobante='+ QuotedStr(pParamatro2), Result, pSucursal)
    else
      _Confecciona_Linea_Sentencias_Faltantes('lin_comprobantes_de_ventas'    , 'SELECT lin_comprobantes_de_ventas.*     FROM lin_comprobantes_de_ventas     WHERE lin_comprobantes_de_ventas.id_comprobante='+ QuotedStr(pParamatro2), Result, pSucursal);
    _Confecciona_Linea_Sentencias_Faltantes('imp_comprobantes_de_ventas'    , 'SELECT imp_comprobantes_de_ventas.*     FROM imp_comprobantes_de_ventas     WHERE imp_comprobantes_de_ventas.id_empresa='+pParamatro1+' AND imp_comprobantes_de_ventas.id_comprobante='+ QuotedStr(pParamatro2), Result, pSucursal);
    _Confecciona_Linea_Sentencias_Faltantes('comprobantes_vta_mov_sdo_favor'  , 'SELECT comprobantes_vta_mov_sdo_favor.*   FROM comprobantes_vta_mov_sdo_favor   WHERE comprobantes_vta_mov_sdo_favor._destino_id_empresa='+pParamatro1+' AND comprobantes_vta_mov_sdo_favor._destino_id_comprobante='+ QuotedStr(pParamatro2), Result, pSucursal);
    _Confecciona_Linea_Sentencias_Faltantes('facturas_notas_debito_presup'  , 'SELECT facturas_notas_debito_presup.*   FROM facturas_notas_debito_presup   WHERE facturas_notas_debito_presup.id_empresa='+pParamatro1+' AND facturas_notas_debito_presup.id_comprobante='+ QuotedStr(pParamatro2), Result, pSucursal);
    _Confecciona_Linea_Sentencias_Faltantes('movimientos_stock'             , 'SELECT movimientos_stock.*              FROM movimientos_stock              WHERE movimientos_stock.id_empresa='+pParamatro1+' AND movimientos_stock._ventas_id_comprobante='+ QuotedStr(pParamatro2), Result, pSucursal);
    _Confecciona_Linea_Sentencias_Faltantes('cupones_entidad_debito'        , 'SELECT cupones_entidad_debito.*         FROM cupones_entidad_debito         WHERE cupones_entidad_debito.id_empresa='+pParamatro1+' AND cupones_entidad_debito.id_comprobante='+ QuotedStr(pParamatro2), Result, pSucursal);
    _Confecciona_Linea_Sentencias_Faltantes('cupones_entidad_convenio'      , 'SELECT cupones_entidad_convenio.*       FROM cupones_entidad_convenio       WHERE cupones_entidad_convenio.id_empresa='+pParamatro1+' AND cupones_entidad_convenio.id_comprobante='+ QuotedStr(pParamatro2), Result, pSucursal);
    _Confecciona_Linea_Sentencias_Faltantes('cupones_tarjeta_empresa'       , 'SELECT cupones_tarjeta_empresa.*        FROM cupones_tarjeta_empresa        WHERE cupones_tarjeta_empresa.id_empresa='+pParamatro1+' AND cupones_tarjeta_empresa.id_comprobante='+ QuotedStr(pParamatro2), Result, pSucursal);
    _Confecciona_Linea_Sentencias_Faltantes('cupones_promocion'             , 'SELECT cupones_promocion.*              FROM cupones_promocion              WHERE cupones_promocion.id_empresa='+pParamatro1+' AND cupones_promocion.id_comprobante='+ QuotedStr(pParamatro2), Result, pSucursal);
    _Confecciona_Linea_Sentencias_Faltantes('preasientos_contables'         , 'SELECT preasientos_contables.*          FROM preasientos_contables          WHERE preasientos_contables.id_empresa='+pParamatro1+' AND preasientos_contables.id_comprobante='+ QuotedStr(pParamatro2), Result, pSucursal);
    _Confecciona_Linea_Sentencias_Faltantes('partidas_preasientos_contables', 'SELECT partidas_preasientos_contables.* FROM partidas_preasientos_contables WHERE partidas_preasientos_contables.nro_preasiento IN (SELECT preasientos_contables.nro_preasiento FROM preasientos_contables WHERE preasientos_contables.id_empresa='+pParamatro1+' AND preasientos_contables.id_comprobante='+ QuotedStr(pParamatro2)+') AND partidas_preasientos_contables.id_sucursal = (SELECT preasientos_contables.id_sucursal FROM preasientos_contables WHERE preasientos_contables.id_empresa= '+pParamatro1+' AND preasientos_contables.id_comprobante='+ QuotedStr(pParamatro2)+')', Result, pSucursal);
    _Confecciona_Linea_Sentencias_Faltantes('vales_de_entrega'              , 'SELECT vales_de_entrega.*               FROM vales_de_entrega               WHERE vales_de_entrega.id_empresa='+pParamatro1+' AND vales_de_entrega._ventas_id_comprobante='+ QuotedStr(pParamatro2), Result, pSucursal);
  end
  else if pTipo_Comprobante= 'envio' then
  begin
    Result:= TStringList.Create;
    _Confecciona_Linea_Sentencias_Faltantes('comprobantes_de_envios'     , 'SELECT comprobantes_de_envios.*     FROM comprobantes_de_envios     WHERE comprobantes_de_envios.id_empresa='+pParamatro1+' AND comprobantes_de_envios.id_comprobante='+ QuotedStr(pParamatro2), Result, pSucursal);
    _Confecciona_Linea_Sentencias_Faltantes('notas_de_envio'             , 'SELECT notas_de_envio.*             FROM notas_de_envio             WHERE notas_de_envio.id_empresa='+pParamatro1+' AND notas_de_envio.id_comprobante='+ QuotedStr(pParamatro2), Result, pSucursal);
      _Confecciona_Linea_Sentencias_Faltantes('lin_comprobantes_de_envios' , 'SELECT lin_comprobantes_de_envios.* FROM lin_comprobantes_de_envios WHERE lin_comprobantes_de_envios.id_comprobante='+ QuotedStr(pParamatro2), Result, pSucursal);
    _Confecciona_Linea_Sentencias_Faltantes('movimientos_stock'             , 'SELECT movimientos_stock.*              FROM movimientos_stock              WHERE movimientos_stock.id_empresa='+pParamatro1+' AND movimientos_stock._envios_id_comprobante='+ QuotedStr(pParamatro2), Result, pSucursal);
  end
  else if pTipo_Comprobante= 'comprobante_compra' then
  begin
    Result:= TStringList.Create;
    _Confecciona_Linea_Sentencias_Faltantes('comprobantes_compras'          , 'SELECT comprobantes_compras.* FROM comprobantes_compras WHERE comprobantes_compras.id_proveedor= '+pParamatro1+' AND comprobantes_compras.id_comprobante = '+ QuotedStr(pParamatro2), Result, pSucursal);
    _Confecciona_Linea_Sentencias_Faltantes('imp_comprobantes_de_compras'   , 'SELECT imp_comprobantes_de_compras.* FROM imp_comprobantes_de_compras WHERE imp_comprobantes_de_compras.id_proveedor= '+pParamatro1+' AND imp_comprobantes_de_compras.id_comprobante = '+ QuotedStr(pParamatro2), Result, pSucursal);
    _Confecciona_Linea_Sentencias_Faltantes('preasientos_contables'         , 'SELECT preasientos_contables.*          FROM preasientos_contables          WHERE preasientos_contables._compras_id_proveedor='+pParamatro1+' AND preasientos_contables._compras_id_comprobante='+ QuotedStr(pParamatro2), Result, pSucursal);
    _Confecciona_Linea_Sentencias_Faltantes('partidas_preasientos_contables', 'SELECT partidas_preasientos_contables.* FROM partidas_preasientos_contables WHERE partidas_preasientos_contables.nro_preasiento IN (SELECT preasientos_contables.nro_preasiento FROM preasientos_contables WHERE preasientos_contables._compras_id_proveedor='+pParamatro1+' AND preasientos_contables._compras_id_comprobante='+ QuotedStr(pParamatro2)+') AND partidas_preasientos_contables.id_sucursal = (SELECT preasientos_contables.id_sucursal FROM preasientos_contables WHERE preasientos_contables._compras_id_proveedor= '+pParamatro1+' AND preasientos_contables._compras_id_comprobante='+ QuotedStr(pParamatro2)+')', Result, pSucursal);
  end
  else if pTipo_Comprobante= 'persona' then
  begin
    Result:= TStringList.Create;
    _Confecciona_Linea_Sentencias_Faltantes('personas'                      , 'SELECT personas.* FROM personas WHERE personas.id_persona= '+pParamatro1, Result, pSucursal);
    _Confecciona_Linea_Sentencias_Faltantes('personas_tel'                  , 'SELECT personas_tel.* FROM personas_tel WHERE personas_tel.id_persona= '+pParamatro1, Result, pSucursal);
    _Confecciona_Linea_Sentencias_Faltantes('personas_email'                , 'SELECT personas_email.* FROM personas_email WHERE personas_email.id_persona= '+pParamatro1, Result, pSucursal);
  end
  else if pTipo_Comprobante= 'registro de cajas' then
  begin
    Result:= TStringList.Create;
    _Confecciona_Linea_Sentencias_Faltantes('registro_de_cajas'             , 'SELECT registro_de_cajas.* FROM registro_de_cajas WHERE registro_de_cajas.id_empresa = '+pParamatro1+' AND registro_de_cajas.id_sucursal = '+pParamatro2+' AND registro_de_cajas.nro_caja = '+pParamatro3+' AND registro_de_cajas.periodo_caja = '+pParamatro4, Result, pSucursal);
  end
  else if pTipo_Comprobante= 'nota de credito' then
  begin
    Result:= TStringList.Create;
    oNota_de_Credito:= _Gestor_Comprobante_Ventas._Buscar_NC(StrToInt(pParamatro1), pParamatro2, oParametro);
    if Assigned(oNota_de_Credito) then
    begin
      _Confecciona_Linea_Sentencias_Faltantes('comprobantes_de_ventas'        , 'SELECT comprobantes_de_ventas.*         FROM comprobantes_de_ventas         WHERE comprobantes_de_ventas.id_empresa='+pParamatro1+' AND comprobantes_de_ventas.id_comprobante='+ QuotedStr(pParamatro2), Result, pSucursal);
      if oParametro.Estructura_Id_empresa_lin_comprobantes_de_ventas then
        _Confecciona_Linea_Sentencias_Faltantes('lin_comprobantes_de_ventas'    , 'SELECT lin_comprobantes_de_ventas.*     FROM lin_comprobantes_de_ventas     WHERE lin_comprobantes_de_ventas.id_empresa='+pParamatro1+' AND lin_comprobantes_de_ventas.id_comprobante='+ QuotedStr(pParamatro2), Result, pSucursal)
      else
        _Confecciona_Linea_Sentencias_Faltantes('lin_comprobantes_de_ventas'    , 'SELECT lin_comprobantes_de_ventas.*     FROM lin_comprobantes_de_ventas     WHERE lin_comprobantes_de_ventas.id_comprobante='+ QuotedStr(pParamatro2), Result, pSucursal);
      _Confecciona_Linea_Sentencias_Faltantes('imp_comprobantes_de_ventas'    , 'SELECT imp_comprobantes_de_ventas.*     FROM imp_comprobantes_de_ventas     WHERE imp_comprobantes_de_ventas.id_empresa='+pParamatro1+' AND imp_comprobantes_de_ventas.id_comprobante='+ QuotedStr(pParamatro2), Result, pSucursal);
      _Confecciona_Linea_Sentencias_Faltantes('notas_credito'                 , 'SELECT notas_credito.*                  FROM notas_credito                  WHERE notas_credito.id_empresa='+pParamatro1+' AND notas_credito._nc_id_comprobante='+ QuotedStr(pParamatro2), Result, pSucursal);
      _Confecciona_Linea_Sentencias_Faltantes('movimientos_stock'             , 'SELECT movimientos_stock.*              FROM movimientos_stock              WHERE movimientos_stock.id_empresa='+pParamatro1+' AND movimientos_stock._ventas_id_comprobante='+ QuotedStr(pParamatro2), Result, pSucursal);
      _Confecciona_Linea_Sentencias_Faltantes('preasientos_contables'         , 'SELECT preasientos_contables.*          FROM preasientos_contables          WHERE preasientos_contables.id_empresa='+pParamatro1+' AND preasientos_contables.id_comprobante='+ QuotedStr(pParamatro2), Result, pSucursal);
      _Confecciona_Linea_Sentencias_Faltantes('partidas_preasientos_contables', 'SELECT partidas_preasientos_contables.* FROM partidas_preasientos_contables WHERE partidas_preasientos_contables.nro_preasiento IN (SELECT preasientos_contables.nro_preasiento FROM preasientos_contables WHERE preasientos_contables.id_empresa='+pParamatro1+' AND preasientos_contables.id_comprobante='+ QuotedStr(pParamatro2)+') AND partidas_preasientos_contables.id_sucursal = (SELECT preasientos_contables.id_sucursal FROM preasientos_contables WHERE preasientos_contables.id_empresa= '+pParamatro1+' AND preasientos_contables.id_comprobante='+ QuotedStr(pParamatro2)+')', Result, pSucursal);

      if _Gestor_Comprobante_Ventas._Saldo_Factura_y_NC(oNota_de_Credito.Id_Empresa, oNota_de_Credito._Fac_Id_Comprobante) <= 0.1 then
      begin
        if (oNota_de_Credito.Fp_Convenio+oNota_de_Credito.Calc_TotalCobertuCIva)>0 then
        begin
          _Confecciona_Linea_Sentencias_Faltantes('cupones_entidad_convenio', 'UPDATE cupones_entidad_convenio SET estado_cupon = '+QuotedStr('AN')+' WHERE id_empresa = '+pParamatro1+' AND id_comprobante = '+Chr(39)+oNota_de_Credito._Fac_Id_Comprobante+Chr(39)+' AND estado_cupon = '+QuotedStr('SP')+';' , Result, pSucursal, 'SENTENCIA MANUAL');
        end;
        if oNota_de_Credito.Fp_Tarjeta_Debito>0 then
        begin
          _Confecciona_Linea_Sentencias_Faltantes('cupones_entidad_debito', 'UPDATE cupones_entidad_debito SET estado_cupon = '+QuotedStr('AN')+' WHERE id_empresa = '+pParamatro1+' AND id_comprobante = '+Chr(39)+oNota_de_Credito._Fac_Id_Comprobante+Chr(39)+' AND estado_cupon = '+QuotedStr('SP')+';' , Result, pSucursal, 'SENTENCIA MANUAL');
        end;
        if oNota_de_Credito.Fp_Cheque>0 then
        begin
          _Confecciona_Linea_Sentencias_Faltantes('cheques_recibidos', 'DELETE FROM cheques_recibidos WHERE id_empresa = '+pParamatro1+' AND id_comprobante = '+Chr(39)+oNota_de_Credito._Fac_Id_Comprobante+Chr(39)+' AND estado_cheque = '+QuotedStr('PE')+';' , Result, pSucursal, 'SENTENCIA MANUAL');
        end;
        if oNota_de_Credito.Fp_Tarjeta_Empresa>0 then
        begin
          _Confecciona_Linea_Sentencias_Faltantes('cupones_tarjeta_empresa', 'UPDATE cupones_tarjeta_empresa SET estado_cupon = '+QuotedStr('AN')+' WHERE id_empresa = '+pParamatro1+' AND id_comprobante = '+Chr(39)+oNota_de_Credito._Fac_Id_Comprobante+Chr(39)+' AND estado_cupon = '+QuotedStr('SP')+';' , Result, pSucursal, 'SENTENCIA MANUAL');
        end;
      end;
      FreeAndNil(oNota_de_Credito);
    end;
  end;
end;


procedure TVSComunicacionParaReplicacion01._Consultar_Cupones_No_Informados_IQVIA(
  pModo: String; pFecha_Desde, pFecha_Hasta: TDate);
var
  oIQVIA                          : TIQVIA;
  oLista_Comprobantes_Sin_Informar: TObjectList;
  oGestor_Comprobante_Ventas      : TGestor_Comprobante_Ventas;
  oFactura_ND_Presupuesto         : TFactura_ND_Presupuesto;
  iItem,iItem2                    : Integer;
  iItemSucursal                   : Integer;
  sRespuesta                      : String;
  sRespuesta_Query                : String;
  oSucursal                       : TSucursal;
  dsSucursales                    : TDataSet;
  oLista_Sucursales               : TObjectList;
  dtFecha                         : TDate;
  //sFecha_Hasta                    : String;
  //I: Integer;
  dsComprobantes                  : TDataSet;

  oGestor_Lanzador_Zafiro: TGestor_Lanzador_Zafiro;

  oGestor_Tipo_Validacion_Receta: TGestor_Tipo_Validacion_Receta;
  oValidadores_Por_Sucursal:TValidadores_Por_Sucursal;
  oTipo_Validacion_Receta  : TTipo_Validacion_Receta;
  sCarpetaErrores: String;
  dsEmpresas: TDataSet;
begin
  // Funcion que envia el oParametro y recupera todos los comprobantes que
  // no fueron informados.

  //pModo = 'ARCHIVO'
  //pModo = 'WS'

  //dmConexion.Conexion.connected:= False;

  //mmIqvia_Log_generacionManual.Lines.Clear;
  _Grabar_LogErrores_Comunicacion(-2, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Inicializando proceso... ');

  oGestor_Lanzador_Zafiro          := TGestor_Lanzador_Zafiro.Create;


  oLista_Sucursales                := TObjectList.Create(True);
  oGestor_Comprobante_Ventas       := TGestor_Comprobante_Ventas.Create();
  dsSucursales                     := _Gestor_Sucursal._Get_Todas_Sucursales;

  // Estructura de control que agrega unicamente las Sucursales que van a informar
  // los Cupones, a la lista.
  if Assigned(dsSucursales) then
  begin
    dsSucursales.First;
    if dsSucursales.RecordCount > 0 then
    begin
      while not dsSucursales.eof do
      begin
        if dsSucursales.FieldByName('informar_iqvia').AsInteger=1 then
        begin
          oSucursal:= TSucursal.Create(dsSucursales.Fields);
          oLista_Sucursales.Add(oSucursal);
        end;
        dsSucursales.Next;
      end;
    end;
    FreeAndNil(dsSucursales);
  end;
  oGestor_Tipo_Validacion_Receta:= TGestor_Tipo_Validacion_Receta.Create;
  oTipo_Validacion_Receta       := oGestor_Tipo_Validacion_Receta._Buscar(5); // Imed - Antes 1

  if Assigned(oGestor_Tipo_Validacion_Receta) then
    FreeAndNil(oGestor_Tipo_Validacion_Receta);

  if Assigned(oTipo_Validacion_Receta) then
  begin
    if oLista_Sucursales.Count > 0 then
    begin

      for iItemSucursal := 0 to oLista_Sucursales.Count - 1 do
      begin

        oValidadores_Por_Sucursal:= oTipo_Validacion_Receta._Get_Validadores_Por_Sucursal(TSucursal(oLista_Sucursales[iItemSucursal]).Id_empresa, TSucursal(oLista_Sucursales[iItemSucursal]).Id_sucursal);
        if Assigned(oValidadores_Por_Sucursal) then
        begin

          _Grabar_LogErrores_Comunicacion(-2, 'Sucursal: '+ IntToStr(TSucursal(oLista_Sucursales[iItemSucursal]).Id_sucursal)+' - '+TSucursal(oLista_Sucursales[iItemSucursal]).Des_sucursal);

          dtFecha := pFecha_Desde;
          while dtFecha<=pFecha_Hasta do
          begin


            oLista_Comprobantes_Sin_Informar := TObjectList.Create(True);

            // Estructura de control que agrega a la lista unicamente los Comprobantes
            // que van a informarse.

            //dsComprobantes:= _Gestor_Lanzador_Zafiro._Buscar_Cupones_Sin_Informar_IQVIA(oParametros, TSucursal(oLista_Sucursales[iItemSucursal]).id_Sucursal, FormatDateTime('yyyy/mm/dd', dtFecha), FormatDateTime('yyyy/mm/dd', dtFecha));
            try
              dsComprobantes:= oGestor_Lanzador_Zafiro._Buscar_Comprobantes_Sin_Informar_IQVIA(TSucursal(oLista_Sucursales[iItemSucursal]).Id_empresa, TSucursal(oLista_Sucursales[iItemSucursal]).id_Sucursal, FormatDateTime('yyyy/mm/dd', dtFecha), FormatDateTime('yyyy/mm/dd', dtFecha));
            except
            end;

            if Assigned(dsComprobantes) then
            begin

              with dsComprobantes do
              begin
                First;
                if RecordCount > 0 then
                begin
                  //mmIqvia_Log_generacionManual.Lines.Add('   Fecha: '+ FormatDateTime('dd/mm/yyyy', dtFecha));

                  //if pmodo= 'ARCHIVO' then
                  //  mmIqvia_Log_generacionManual.Lines.Add('   Fecha: '+ FormatDateTime('dd/mm/yyyy', dtFecha)+' - Creando Archivos para '+ IntToStr(RecordCount)+' Comprobantes...')
                  //else
                    _Grabar_LogErrores_Comunicacion(-2, '   Fecha: '+ FormatDateTime('dd/mm/yyyy', dtFecha)+' - Informando '+ IntToStr(RecordCount)+' Comprobantes...');

                  // Comento la estructura de control While porque la consulta SQL recupera
                  // muchos registros, por lo tanto utilizo una estructura de control
                  // For, que recupere los primeros 19 registros.
                  while not eof do
                  begin
                    oFactura_ND_Presupuesto:= Nil;
                    try
                      oFactura_ND_Presupuesto := oGestor_Comprobante_Ventas._Buscar_FAC_ND_PRE(FieldByName('id_empresa').AsInteger, FieldByName('id_comprobante').AsString, oParametro);
                    except
                      oFactura_ND_Presupuesto:= Nil;
                    end;
                    if Assigned(oFactura_ND_Presupuesto) then
                    begin
                      oLista_Comprobantes_Sin_Informar.Add(oFactura_ND_Presupuesto);
                    end;

                    //edtid_Comprobante.Text         := FieldByName('id_comprobante').AsString;
                    Next;
                  end;
                end;
                Close;
              end;


              sCarpetaErrores := GetCurrentDir;
              if not DirectoryExists('IQVIA_Errores') then
                CreateDir('IQVIA_Errores');
              sCarpetaErrores := sCarpetaErrores + '\'+'IQVIA_Errores';

              // Bloque de sentencias que envía la Lista que contiene los comprobantes
              // sin informar.
              oIQVIA     := TIQVIA.Create(oLista_Comprobantes_Sin_Informar, oParametro, oValidadores_Por_Sucursal.Adesfa_Codigo_Prestador, sCarpetaErrores);

              try
                sRespuesta := oIQVIA._GenerarSolicitud_IQVIA(FormatDateTime('yyyymmdd', dtFecha), TSucursal(oLista_Sucursales[iItemSucursal]).Id_empresa, TSucursal(oLista_Sucursales[iItemSucursal]).Id_sucursal, False, pModo);

              except
                on E: Exception do
                begin
                  sRespuesta := 'Error al procesar "_GenerarSolicitud_IQVIA" -  Error: '+ (E.Message);
                end;
              end;

              if Assigned(oIQVIA) then
                FreeAndNil(oIQVIA);

              if Assigned(oLista_Comprobantes_Sin_Informar) then
                FreeAndNil(oLista_Comprobantes_Sin_Informar);

              if Length(Trim(sRespuesta))>0  then
              begin
                _Grabar_LogErrores_Comunicacion(-2, sRespuesta);
                //mmIqvia_Log_generacionManual.Lines.Add('Proceso Interrumpido!');

                //if Assigned(oGestor_Comprobante_Ventas) then
                //  FreeAndNil(oGestor_Comprobante_Ventas);
                //if Assigned(oGestor_Sucursales) then
                //  FreeAndNil(oGestor_Sucursales);
                //if Assigned(oLista_Sucursales) then
                //  FreeAndNil(oLista_Sucursales);
                //if Assigned(oParametros) then
                //  FreeAndNil(oParametros);

                //dmConexion.Conexion.connected:= False;

                //Exit
              end;

            end;
            dtFecha := dtFecha + 1;
          end; // while dtFecha<=pFecha_Hasta do

          // Libero oValidadores_Por_Sucursal
          if Assigned(oValidadores_Por_Sucursal) then
            FreeAndNil(oValidadores_Por_Sucursal);

        end;
      end;
      _Grabar_LogErrores_Comunicacion(-2, 'Proceso Finalizado!');
    end;

  end;

  if Assigned(oGestor_Lanzador_Zafiro) then
    FreeAndNil(oGestor_Lanzador_Zafiro);
  if Assigned(oGestor_Comprobante_Ventas) then
    FreeAndNil(oGestor_Comprobante_Ventas);
  if Assigned(oLista_Sucursales) then
    FreeAndNil(oLista_Sucursales);
  if Assigned(oTipo_Validacion_Receta) then
    FreeAndNil(oTipo_Validacion_Receta);

//  dmConexion.Conexion.connected:= False;

end;

function TVSComunicacionParaReplicacion01._Actualiza_Ultimo_Log_Procesado(pId_ProcesoActualizacion: Smallint): Boolean;
var
  sCadena_error: String;
begin
  //Actualizo el ultimo id_archivo_log_procesado y ultimo el pos_log_procesado es decir pongo el limite superior
  //que me indica la ultima transacción de la BD origen a la cual se tienen que empardar las demas BD
  Result:= False;
  try
    //
    qrySelect_LastBinaryLogOrigen.Close;
    qrySelect_LastBinaryLogOrigen.Open;

    sUltimo_archivo_log_BDorigen:= qrySelect_LastBinaryLogOrigen.FieldByName('file').AsString;
    iUltimo_Pos_log_BDorigen    := qrySelect_LastBinaryLogOrigen.FieldByName('position').AsInteger;

    qryUpdate_FileandPos_BDorigen.Parameters.ParamByName('pId_archivo_log_procesado').Value  := sUltimo_archivo_log_BDorigen;
    qryUpdate_FileandPos_BDorigen.Parameters.ParamByName('pPos_log_procesado').Value         := iUltimo_Pos_log_BDorigen;
    qryUpdate_FileandPos_BDorigen.Parameters.ParamByName('pId_sucursal').Value               := oParametro.Comu_Id_Sucursal;

    qryUpdate_FileandPos_BDorigen.Connection.BeginTrans;
    qryUpdate_FileandPos_BDorigen.ExecSQL;
    qryUpdate_FileandPos_BDorigen.Connection.CommitTrans;
    Result:= True;

//    sCadena_error:= DateTimeToStr(Now)+' 99---Sucursal:'+IntToStr(oParametro.Comu_Id_Sucursal)+' - '+'Sucursal Local'+' -  Actualizo: '+ sUltimo_archivo_log_BDorigen+' / '+IntToStr(iUltimo_Pos_log_BDorigen) + '  /  Proceso: '+IntToStr(pId_ProcesoActualizacion);
//    _Grabar_LogErrores_Comunicacion(oParametro.Comu_Id_Sucursal, sCadena_error);


  Except
    on E: Exception do
    begin

      sCadena_error:= DateTimeToStr(Now)+'  Error al actualizar el úlimo log procesado en la DB Local -  Error: '+ (E.Message);
      Result:= False;
      if qryUpdate_FileandPos_BDorigen.Connection.InTransaction = true then
        qryUpdate_FileandPos_BDorigen.Connection.RollbackTrans;
      _Grabar_LogErrores_Comunicacion(pId_ProcesoActualizacion, sCadena_error);
    end;
  end;
end;


function TVSComunicacionParaReplicacion01._ServiceStop(pMachine, pServiceName: String): Boolean;
// aMachine is UNC path or local machine if empty
var
  h_manager, h_svc: SC_Handle;
  ServiceStatus: TServiceStatus;
  dwCheckPoint: DWORD;
begin

  h_manager := OpenSCManager(PChar(pMachine), nil, SC_MANAGER_CONNECT);
  if h_manager > 0 then
  begin
    h_svc := OpenService(h_manager, PChar(pServiceName), SERVICE_STOP or SERVICE_QUERY_STATUS);
    if h_svc > 0 then
    begin
      if (ControlService(h_svc, SERVICE_CONTROL_STOP, ServiceStatus)) then
      begin
        if (QueryServiceStatus(h_svc, ServiceStatus)) then
        begin
          while (SERVICE_STOPPED <> ServiceStatus.dwCurrentState) do
          begin
            dwCheckPoint := ServiceStatus.dwCheckPoint;
            Sleep(ServiceStatus.dwWaitHint);
            if (not QueryServiceStatus(h_svc, ServiceStatus)) then
              // couldn't check status
              break;
            if (ServiceStatus.dwCheckPoint < dwCheckPoint) then
              Break;
          end;
        end;
      end;
      CloseServiceHandle(h_svc);
    end;
    CloseServiceHandle(h_manager);
  end;

  Result := (SERVICE_STOPPED = ServiceStatus.dwCurrentState);

end;


function TVSComunicacionParaReplicacion01._ObtenerNivele_Rotacion(pId_Empresa: Integer; pId_sucursal: Integer; pId_articulo: String; pFecha_desde: TDate; pFecha_hasta: TDate):TNiveles_de_Rotacion;
var
  dsVentas_Anteriores    : TDataSet;
  dsNivel_Rotacion       : TDataSet;
begin
  Result:=Nil;
  dsVentas_Anteriores := _Gestor_Articulos._Buscar_Ventas_Anteriores_x_ArtSuc(pId_Empresa, pId_sucursal, pId_articulo, pFecha_desde, pFecha_hasta);
  if Assigned(dsVentas_Anteriores) then
    dsNivel_Rotacion    := _Gestor_Niveles_Rotacion._Buscar_x_VtasAnteriores(dsVentas_Anteriores.FieldByName('cantidad').AsInteger)
  else
    dsNivel_Rotacion    := _Gestor_Niveles_Rotacion._Buscar_x_VtasAnteriores(0);

  if Assigned(dsNivel_Rotacion) then
  begin
    if dsNivel_Rotacion.RecordCount > 0 then
    begin
      Result:= TNiveles_de_Rotacion.Create(dsNivel_Rotacion.Fields);
    end;
  end;
end;

procedure TVSComunicacionParaReplicacion01._Subir_Comp_Pend_Servidor_FTP;
var
  dsComprobantesPendientes: TDataSet;
  sNombreArchivo: string;
  bCargadoEnFTP: Boolean;
  sMensaje, sError : string;
  sURL: string;
begin
  dsComprobantesPendientes := nil;
  dsComprobantesPendientes := _Gestor_log_subida_Comp._Archivos_Pendientes_de_Carga;
  if Assigned(dsComprobantesPendientes) then
  begin
    if dsComprobantesPendientes.RecordCount > 0 then
    begin
      with oParametro do
      try
        _dm_Comrobante_Ventas._Conectar_Servidor_FTP(Fact_Online_FTP_Host, Fact_Online_FTP_User, Fact_Online_FTP_Pass, Fact_Online_FTP_Port, 3000);
      except
        bCargadoEnFTP := False;
      end;

      dsComprobantesPendientes.First;
      while not(dsComprobantesPendientes.Eof) do
      begin
        sError   := EmptyStr;
        if FileExists(dsComprobantesPendientes.FieldByName('observaciones').AsString) then // el Archivo se guardó, subo el pdf a la nube
        begin
          if _dm_Comrobante_Ventas.IdFTP1.Connected then
          begin
            sNombreArchivo := ExtractFileName(dsComprobantesPendientes.FieldByName('observaciones').AsString);
            bCargadoEnFTP := _dm_Comrobante_Ventas._Cargar_Archivo_En_Servidor_FTP(dsComprobantesPendientes.FieldByName('observaciones').AsString, sNombreArchivo);

            if bCargadoEnFTP then
            begin
              sMensaje := _Gestor_log_subida_Comp._Indicar_Archivo_Subido(dsComprobantesPendientes.FieldByName('id_empresa').AsInteger, dsComprobantesPendientes.FieldByName('id_comprobante').AsString);
              if sMensaje = '' then
              begin
                sURL := oParametro.Fact_Online_FTP_URL_base + sNombreArchivo;
                sMensaje := _Gestor_Comprobante_Ventas._Insertar_URL_en_CV(dsComprobantesPendientes.FieldByName('id_empresa').AsInteger, dsComprobantesPendientes.FieldByName('id_comprobante').AsString, sURL);
                if sMensaje = '' then
                begin
                  sError := ' - El archivo ' + dsComprobantesPendientes.FieldByName('observaciones').AsString + ' se cargó en el servidor FTP' +
                            ' - Se registró con éxito en la tabla de hechos el log de subida de comprobante' +
                            ' - Se guardó la Dirección URL ' + sURL + ' correspondiente al comprobante';
                end
                else
                begin
                  sError := ' - El archivo ' + dsComprobantesPendientes.FieldByName('observaciones').AsString + ' se cargó en el servidor FTP' +
                              ' - Se registró con éxito en la tabla de hechos el log de subida de comprobante' +
                              ' - No se pudo guardar URL en comprobante por error : ' + sMensaje;
                end;
              end
              else
              begin
                sError := ' - El archivo ' + dsComprobantesPendientes.FieldByName('observaciones').AsString + ' se cargó en el servidor FTP' +
                            ' - No se pudo registrar en la tabla de hechos el log de subida de comprobante por error : ' + sMensaje;
              end;
            end
            else
            begin
              sError := ' - Falló al cargar el archivo ' + dsComprobantesPendientes.FieldByName('observaciones').AsString + ' en el servidor FTP' +
                        ' - Revise la conexión a Internet y/o los Parámetros del Sistema que gestionan esta conexión';
            end;
          end
          else
          begin
            sError := ' - Falló al cargar el archivo ' + dsComprobantesPendientes.FieldByName('observaciones').AsString + ' en el servidor FTP' +
                      ' - Revise la conexión a Internet y/o los Parámetros del Sistema que gestionan esta conexión';
          end;
        end
        else
        begin
          sError := ' - El archivo ' + dsComprobantesPendientes.FieldByName('observaciones').AsString + ' no existe.';
        end;

        if Length(sError) > 0 then
          _Grabar_LogErrores_Comunicacion(-6, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ sError);

        dsComprobantesPendientes.Next;
      end;
      _dm_Comrobante_Ventas._Desconectar_Servidor_FTP;
    end;
  end;
end;

function TVSComunicacionParaReplicacion01._Tabla_Exceptuada(pTabla: String;
  pSucursal: TSucursal): Boolean;
begin
  Result:= False;
  if Assigned(pSucursal.Tablas_exceptuadas_sucursal) then
    if (pSucursal.Tablas_exceptuadas_sucursal.IndexOf(pTabla)> -1) then
      Result:= True;
  if ((Trim(pTabla) = 'parametros') or (Trim(pTabla) = 'vsmodulo') or (Trim(pTabla) = 'vsimagenes') or (Trim(pTabla) = 'vsmodulobi')
      or (Trim(pTabla) = 'comunicacion') or (Trim(pTabla) = 'log_comunicacion_externo') or (Trim(pTabla) = 'sucursales_comunicacion') or  (Trim(pTabla) = 'sucursales_log_envio') or (Trim(pTabla) = 'tablas_exceptuadas_sucursal')
      or (Trim(pTabla) = 'inv_articulos_cb') or (Trim(pTabla) = 'inv_auditoria') or (Trim(pTabla) = 'inv_incidencias') or (Trim(pTabla) = 'inv_relevamientos') or (Trim(pTabla) = 'inv_relevamiento_filtro') or (Trim(pTabla) = 'inv_relevamiento_usuario') or (Trim(pTabla) = 'inv_parametros')) then
    Result:= True;
end;

function TVSComunicacionParaReplicacion01._Transmision_de_Sentencias_Por_Comprobante(pSucursal: TSucursal) : String;
var
  dtFechaHoraActual: TDateTime;
  sRespuesta : String;
  dtFechaHoraUltimaActualizacion: TDateTime;

  dtHoraProximaActualizacion: TDateTime;
  dtFechaHoraProximaActualizacion: TDateTime;


begin

  // Saco la Hora actual de la BD
  if qrySelect_Now.Active then
    qrySelect_Now.Close;
  qrySelect_Now.Open;
  if qrySelect_Now.RecordCount > 0 then
    dtFechaHoraActual:= qrySelect_Now.FieldByName('Fecha_Hora').AsDateTime
  else
    dtFechaHoraActual:= Now;
  qrySelect_Now.Close;
  //

  dtFechaHoraUltimaActualizacion:= pSucursal.Comu_Fecha_Hora_Ultima_Revision;

//  dtFechaHoraUltimaActualizacion:= StrToDateTime('26/04/2023 02:00:00');
//  dtFechaHoraActual             := StrToDateTime('27/04/2023 03:01:00');

  if (dtFechaHoraActual - dtFechaHoraUltimaActualizacion)>1  then
  begin
    try
      //dtHoraProximaActualizacion:= StrToTime(LeftStr(oParametro.Prec_Hora_Actualizacion,5));
      dtHoraProximaActualizacion:= StrToTime(LeftStr('01:00',5));
    except
      dtHoraProximaActualizacion:= 0;
    end;
    dtFechaHoraUltimaActualizacion:= DateOf(dtFechaHoraActual) - 1 + dtHoraProximaActualizacion;
  end;


  if dtFechaHoraUltimaActualizacion <> 0  then
  begin
    try
      dtHoraProximaActualizacion:= StrToTime(LeftStr(oParametro.Prec_Hora_Actualizacion,5));
    except
      dtHoraProximaActualizacion:= 0;
    end;
    dtFechaHoraProximaActualizacion:= DateOf(dtFechaHoraUltimaActualizacion) + 1 + dtHoraProximaActualizacion;
  end;

  if dtFechaHoraActual > dtFechaHoraProximaActualizacion then
  begin
    _Grabar_LogErrores_Comunicacion(-8, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Ejecutando Transmisión de Sentencias por Comprobante...', pSucursal.Id_Sucursal);

    sRespuesta:= _Ejecuta_Transmision_de_Sentencias_Por_Comprobante(pSucursal);

    if Length(sRespuesta)=0 then
    begin
      // Actualizar Prec_Fecha_Hora_Ultima_Actualizacion
      _Modificar_Comu_Fecha_Hora_Ultima_Revision(pSucursal.Id_Sucursal);
      pSucursal.Comu_Fecha_Hora_Ultima_Revision := Now();
      //
    end;
  end;

end;

function TVSComunicacionParaReplicacion01._Ejecuta_Transmision_de_Sentencias_Por_Comprobante(
  pSucursal: TSucursal): String;
var
  dFecha_Desde: TDate;
  dFecha_Hasta: TDate;
  oGestor_Comunicacion_Manual: TGestor_Comunicacion_Manual;
  dtFecha: TDate;
  dsComprobantes: TDataSet;
  dsComprobantesRemotos: TDataSet;
  iTipoComp: SmallInt;
  sTipo_Comprobante: String;

  bConecto: Boolean;
  bEncontroRegistroBDRemota: Boolean;
  slSentencias: TStringList;
  sRespuesta: string;
begin
  {
  pedido
  envio domicilio
  factura
  nota de credito
  presentacion ED
  presentacion EC
  envio
  envio domicilio
  ajuste stock
  ingreso stock
  vale entrega - ingreso
  }

  Result:= '';

  dFecha_Hasta:= Date - 1;
  dFecha_Desde:= dFecha_Hasta - 35;

  oGestor_Comunicacion_Manual := TGestor_Comunicacion_Manual.Create;

  if Assigned(pSucursal) then
  begin
    // ojo pv
    //for iTipoComp := 1 to 12 do
    for iTipoComp := 2 to 14 do   // Para que no actualice Personas
    begin
      Case iTipoComp of
        1 : sTipo_Comprobante := 'persona';
        2 : sTipo_Comprobante := 'pedido distribuido';
        3 : sTipo_Comprobante := 'pedido';
        4 : sTipo_Comprobante := 'envio domicilio' ;
        5 : sTipo_Comprobante := 'factura' ;
        6 : sTipo_Comprobante := 'nota de credito' ;
        7 : sTipo_Comprobante := 'ajuste stock' ;
        8 : sTipo_Comprobante := 'solicitud NC' ;
        9 : sTipo_Comprobante := 'remito' ;
        10 : sTipo_Comprobante := 'ingreso stock' ;
        11 : sTipo_Comprobante := 'envio' ;
        12 : sTipo_Comprobante := 'presentacion ED' ;
        13 : sTipo_Comprobante := 'presentacion EC' ;
        14 : sTipo_Comprobante := 'vale entrega - ingreso';
      end;

      _Grabar_LogErrores_Comunicacion(-8, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - '+sTipo_Comprobante+': Proceso Iniciado!', pSucursal.Id_Sucursal);

      dtFecha := dFecha_Desde;
      while dtFecha <= dFecha_Hasta do
      begin
        _Grabar_LogErrores_Comunicacion(-8, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Inicializando...', pSucursal.Id_Sucursal);
        _Grabar_LogErrores_Comunicacion(-8, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Fecha: ' + FormatDateTime('dd/mm/yyyy', dtFecha), pSucursal.Id_Sucursal);

        dsComprobantes := oGestor_Comunicacion_Manual._Buscar_Comprobantes_A_Verificar(sTipo_Comprobante, FormatDateTime('yyyy/mm/dd', dtFecha), FormatDateTime('yyyy/mm/dd', dtFecha));

        //try
        //  iId_Sucursal := StrToInt(vsedtbId_Sucursal2.Text);
        //except
        //  iId_Sucursal := 0;
        //end;

        if Assigned(dsComprobantes) then
        begin
          with dsComprobantes do
          begin

            //for iItem := 0 to oSucursales.Count - 1 do
            //begin
              //bTransmiteASuc := False;
              //if iId_Sucursal > 0 then
              //begin
              //  if TSucursal(oSucursales.Items[iItem]).Id_sucursal = iId_Sucursal then
              //    bTransmiteASuc := True;
              //end
              //else
              //begin
              //  bTransmiteASuc := True;
              //end;

              //if bTransmiteASuc then
              //begin
                bConecto := oGestor_Comunicacion_Manual._Conectar_Conexion_Remota(pSucursal.Cadena_conexion);
                if bConecto then
                begin

                  dsComprobantesRemotos := oGestor_Comunicacion_Manual._Buscar_Comprobantes_Conexion_Remota(sTipo_Comprobante, FormatDateTime('yyyy/mm/dd', dtFecha), FormatDateTime('yyyy/mm/dd', dtFecha));

                  if Assigned(dsComprobantesRemotos) then
                  begin
                    if RecordCount > 0 then
                    begin
                      _Grabar_LogErrores_Comunicacion(-8, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Transmitiendo a la sucursal: ' + IntToStr(pSucursal.Id_sucursal) + ' - ' + pSucursal.Des_sucursal, pSucursal.Id_Sucursal);

                      First;
                      if ((sTipo_Comprobante = 'factura') or (sTipo_Comprobante = 'nota de credito') or (sTipo_Comprobante = 'envio') or (sTipo_Comprobante = 'envio domicilio') or (sTipo_Comprobante = 'ajuste stock')  or (sTipo_Comprobante = 'ingreso stock') or (sTipo_Comprobante = 'solicitud NC') or (sTipo_Comprobante = 'remito') or (sTipo_Comprobante = 'pedido distribuido') or (sTipo_Comprobante = 'pedido') or (sTipo_Comprobante = 'persona') or (sTipo_Comprobante = 'vale entrega - ingreso')) then
                      begin

                        while not eof do
                        begin
                          // Aqui verifico si existe comprobante en destino
                          //sRespuesta:=oGestor_Comunicacion_Manual._Veridicar_Si_Existe_Comprobante(cmbTipo_ComprobanteTM.Text, dsComprobantes.Fields[0].AsString, dsComprobantes.Fields[1].AsString);

                          bEncontroRegistroBDRemota := dsComprobantesRemotos.Locate('id_empresa;id', VarArrayOf([dsComprobantes.Fields[0].AsInteger, dsComprobantes.Fields[1].AsString]), [loCaseInsensitive]);

                          //if sRespuesta='No' then
                          if bEncontroRegistroBDRemota = False then
                          begin
                            // Transminir
                            slSentencias := oGestor_Comunicacion_Manual._Confecciona_Sentencias_Faltantes(oParametro.Estructura_Id_empresa_lin_comprobantes_de_ventas, sTipo_Comprobante, dsComprobantes.Fields[0].AsString, dsComprobantes.Fields[1].AsString, pSucursal);

                            if Assigned(slSentencias) then
                            begin
                              if slSentencias.Count > 0 then
                              begin
                                sRespuesta := oGestor_Comunicacion_Manual._Ejecutar_Sentencias_Conexion_Remota(slSentencias);
                                if LeftStr(sRespuesta, 5) = 'Error' then
                                begin
                                  _Grabar_LogErrores_Comunicacion(-8, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Empresa:' + dsComprobantes.Fields[0].AsString + ' - Id:' + dsComprobantes.Fields[1].AsString + ' - ' + sRespuesta, pSucursal.Id_Sucursal);
                                end
                                else
                                begin
                                  _Grabar_LogErrores_Comunicacion(-8, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Empresa:' + dsComprobantes.Fields[0].AsString + ' - Id:' + dsComprobantes.Fields[1].AsString + ' - ' + 'Insertado', pSucursal.Id_Sucursal);
                                end;
                              end;
                              FreeAndNil(slSentencias);
                            end
                            else
                            begin
                            //  mmLog.Lines.Add('Empresa:'+dsComprobantes.Fields[0].AsString+' - Id:'+dsComprobantes.Fields[1].AsString + ' - '+ 'No se generaron sentencias');
                            end;
                            //
                          end                          //else if LeftStr(sRespuesta,5)='Error' then
                          //begin
                          //  mmLog.Lines.Add('Empresa:'+dsComprobantes.Fields[0].AsString+' - Id:'+dsComprobantes.Fields[1].AsString + ' - '+sRespuesta);
                          //end
                          else
                          begin
                            // Si existe
                          end;
                          Next;
                        end; // while not eof do
                      end
                      else if ((sTipo_Comprobante = 'presentacion ED') or (sTipo_Comprobante = 'presentacion EC')) then
                      begin

                        while not eof do
                        begin
                          // Aqui verifico si existe comprobante en destino
                          //sRespuesta:=oGestor_Comunicacion_Manual._Veridicar_Si_Existe_Comprobante(cmbTipo_ComprobanteTM.Text, dsComprobantes.Fields[0].AsString, dsComprobantes.Fields[1].AsString);

                          bEncontroRegistroBDRemota := dsComprobantesRemotos.Locate('id_empresa;nro_presentacion', VarArrayOf([dsComprobantes.Fields[0].AsInteger, dsComprobantes.Fields[1].AsInteger]), [loCaseInsensitive]);

                          //if sRespuesta='No' then
                          if bEncontroRegistroBDRemota = False then
                          begin
                            // Transminir
                            slSentencias := oGestor_Comunicacion_Manual._Confecciona_Sentencias_Faltantes(oParametro.Estructura_Id_empresa_lin_comprobantes_de_ventas, sTipo_Comprobante, dsComprobantes.Fields[0].AsString, dsComprobantes.Fields[1].AsString, pSucursal);
                            if Assigned(slSentencias) then
                            begin
                              if slSentencias.Count > 0 then
                              begin
                                sRespuesta := oGestor_Comunicacion_Manual._Ejecutar_Sentencias_Conexion_Remota(slSentencias);
                                if LeftStr(sRespuesta, 5) = 'Error' then
                                begin
                                  _Grabar_LogErrores_Comunicacion(-8, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Empresa:' + dsComprobantes.Fields[0].AsString + ' - Nro.' + sTipo_Comprobante + ':' + dsComprobantes.Fields[1].AsString + ' - ' + sRespuesta + Chr(13) + 'Sentencias: ' + slSentencias.Text, pSucursal.Id_Sucursal);
                                end
                                else
                                begin
                                  _Grabar_LogErrores_Comunicacion(-8, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Empresa:' + dsComprobantes.Fields[0].AsString + ' - Nro.' + sTipo_Comprobante + ':' + dsComprobantes.Fields[1].AsString + ' - ' + 'Insertado', pSucursal.Id_Sucursal);
                                end;
                              end;
                              FreeAndNil(slSentencias);
                            end
                            else
                            begin
                            //  mmLog.Lines.Add('Empresa:'+dsComprobantes.Fields[0].AsString+' - Id:'+dsComprobantes.Fields[1].AsString + ' - '+ 'No se generaron sentencias');
                            end;
                            //
                          end                          //else if LeftStr(sRespuesta,5)='Error' then
                          //begin
                          //  mmLog.Lines.Add('Empresa:'+dsComprobantes.Fields[0].AsString+' - Id:'+dsComprobantes.Fields[1].AsString + ' - '+sRespuesta);
                          //end
                          else
                          begin
                            // Si existe
                          end;
                          Next;
                        end; // while not eof do
                      end;
                    end; // if RecordCount > 0 then
                  end; // if Assigned(dsComprobantesRemotos) then
                  oGestor_Comunicacion_Manual._Desconectar_Conexion_Remota;
                end
                else
                begin
                  _Grabar_LogErrores_Comunicacion(-8, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - No se pudo conectar a la sucursal: ' + IntToStr(pSucursal.Id_sucursal) + ' - ' + pSucursal.Des_sucursal, pSucursal.Id_Sucursal);
                  Result:= 'No se pudo conectar a la sucursal: ' + IntToStr(pSucursal.Id_sucursal) + ' - ' + pSucursal.Des_sucursal;
                end;

                oGestor_Comunicacion_Manual._Desconectar_Conexion_Remota;
              //end; // if bTransmiteASuc then
            //end; // for iItem := 0 to oSucursales.Count-1 do
            Close;
          end;
        end;
        dtFecha := dtFecha + 1;

        if sTipo_Comprobante = 'persona' then // Sale porque no es por rango de fechas. Debe hacerlo una sola vez
          break;
      end;
      _Grabar_LogErrores_Comunicacion(-8, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - '+sTipo_Comprobante+': Proceso Finalizado! '+Chr(13)+'--------------------------------------------', pSucursal.Id_Sucursal);
      _Grabar_LogErrores_Comunicacion(-8, '', pSucursal.Id_Sucursal);
    end;  // for iTipoComp := 1 to 9 do
  end;
  if Assigned(oGestor_Comunicacion_Manual) then
    FreeAndNil(oGestor_Comunicacion_Manual);
end;

procedure TVSComunicacionParaReplicacion01._ServiceStart(pServiceName: String);
var
  SrvControlManager,
  Srv_Handles : SC_Handle;
  Srv_Status  : TServiceStatus;
  // temp char pointer
  psTemp : PChar;
  // check point
  dwChkP : DWord;
begin

  SrvControlManager := OpenSCManager('', Nil,  SC_MANAGER_CONNECT);
  if(SrvControlManager > 0)then
  begin
    // open a handle to the specified service
    Srv_Handles := OpenService( SrvControlManager, PChar(pServiceName), SERVICE_START or SERVICE_QUERY_STATUS);

    if(Srv_Handles > 0)then
    begin
      psTemp := Nil;
      if(StartService(Srv_Handles, 0, psTemp))then
      begin
        // check status
        if(QueryServiceStatus(Srv_Handles, Srv_Status))then
        begin
          while(SERVICE_RUNNING <> Srv_Status.dwCurrentState)do
          begin
            // dwCheckPoint contains a value that the service increments periodically to report its progress during a lengthy operation.
            // save current value
            dwChkP := Srv_Status.dwCheckPoint;

            // wait a bit before checking status again. dwWaitHint is the estimated amount of time the calling program should wait before calling
            // QueryServiceStatus() again

            // idle events should be
            // handled here...
            Sleep(Srv_Status.dwWaitHint);

            if(not QueryServiceStatus( Srv_Handles, Srv_Status))then
            begin
              // couldn't check status break from the loop
              break;
            end;

            if(Srv_Status.dwCheckPoint < dwChkP)then
            begin
              // QueryServiceStatus didn't increment dwCheckPoint as it should have.
              // avoid an infinite loop by breaking
              break;
            end;
          end;
        end;
      end;

      // close service handle
      CloseServiceHandle(Srv_Handles);
    end;

    // close service control manager handle
    CloseServiceHandle(SrvControlManager);
  end;

end;

procedure TVSComunicacionParaReplicacion01.ppmParar_servicioClick(Sender: TObject);
begin
  _ServiceStop('', Self.Name );
  ReportStatus;
//  CoolTrayIcon1.IconIndex:= 0;

  Timer1.Enabled:= False;
  Timer2.Enabled:= False;
  Timer3.Enabled:= False;
  Timer4.Enabled:= False;
  Timer5.Enabled:= False;
  Timer6.Enabled:= False;
  Timer7.Enabled:= False;
  Timer8.Enabled:= False;
  Timer9.Enabled:= False;
  Timer10.Enabled:= False;
  Timer11.Enabled:= False;
  Timer12.Enabled:= False;
  Timer13.Enabled:= False;
  Timer14.Enabled:= False;
  Timer15.Enabled:= False;
  Timer16.Enabled:= False;
  Timer17.Enabled:= False;
  Timer18.Enabled:= False;
  Timer19.Enabled:= False;
  Timer20.Enabled:= False;
  Timer21.Enabled:= False;
  Timer22.Enabled:= False;
  Timer23.Enabled:= False;
  Timer24.Enabled:= False;
  Timer25.Enabled:= False;
  Timer26.Enabled:= False;
  Timer27.Enabled:= False;
  Timer28.Enabled:= False;
  Timer100.Enabled:= False;

  TimerEtiquetas.Enabled:= False;
end;



procedure TVSComunicacionParaReplicacion01.ServiceDestroy(Sender: TObject);
begin
  if Assigned(oSucursales) then
    FreeAndNil(oSucursales);

  if Assigned(oParametro) then
    FreeAndNil(oParametro);

  if Assigned(oParametros_Formato) then
    FreeAndNil(oParametros_Formato);

  if Assigned(_Gestor_Empresa) then
    FreeAndNil(_Gestor_Empresa);

  if Assigned(_Gestor_Sucursal) then
    FreeAndNil(_Gestor_Sucursal);

  if Assigned(_Gestor_Log_Informe_Automatico) then
    FreeAndNil(_Gestor_Log_Informe_Automatico);

  if Assigned(_Gestor_log_subida_Comp) then
    FreeAndNil(_Gestor_log_subida_Comp);

  if Assigned(_Gestor_Comprobante_Ventas) then
    FreeAndNil(_Gestor_Comprobante_Ventas);

  if Assigned(_dm_Comrobante_Ventas) then
    FreeAndNil(_dm_Comrobante_Ventas);

  if Assigned(_Gestor_Articulos) then
    FreeAndNil(_Gestor_Articulos);

  if Assigned(_Gestor_Niveles_Rotacion) then
    FreeAndNil(_Gestor_Niveles_Rotacion);

  Conexion.Close;
  Conexion_Remota.Close;
end;

procedure TVSComunicacionParaReplicacion01.ppmReiniciar_servicioClick(Sender: TObject);
begin
  _ServiceStop('', Self.Name );
  ReportStatus;
//  CoolTrayIcon1.IconIndex:= 0;
  _ServiceStart(Self.Name);
  ReportStatus;
//  CoolTrayIcon1.IconIndex:= 1;
  _Hab_des_Timer(Self.Tag, True);
end;


procedure TVSComunicacionParaReplicacion01.StopService1Click(
  Sender: TObject);
begin
  frmAcercaDe := TfrmAcercaDe.create(self);
  frmAcercaDe.ShowModal;
  FreeAndNil(frmAcercaDe);
end;

procedure TVSComunicacionParaReplicacion01._CargarItemsMenu;
var
  i: Smallint;
  subMenu: TMenuItem;
begin

  for i:=(PopupMenu_TrayServicio.Items.Count -1) downto 5 do
  begin
    PopupMenu_TrayServicio.Items[i].Free;
  end;

  if Assigned(oSucursales) then
  begin
    if Assigned(oParametro) then
    begin
      for i:=0 to oSucursales.Count-1 do
      begin
        if oParametro.Comu_Id_Sucursal <> TSucursal(oSucursales.Items[i]).Id_sucursal then
        begin
          subMenu:= TMenuItem.Create(PopupMenu_TrayServicio);
          subMenu.Caption:= IntToStr(TSucursal(oSucursales.Items[i]).Id_sucursal)+' - '+TSucursal(oSucursales.Items[i]).Des_sucursal;
          if bErrorComunicacionConexionLocal then
            subMenu.Caption:= subMenu.Caption+ ' -> Error de Conexión BD Local'
          else
            if TSucursal(oSucursales.Items[i]).GraboError then
              subMenu.Caption:= subMenu.Caption+ ' -> No Transmitiendo'
            else
              subMenu.Caption:= subMenu.Caption+ ' -> Transmitiendo';
          PopupMenu_TrayServicio.Items.add(subMenu);
        end;
      end;//end for i:=0 to oSucursales.Count-1 do
    end;
  end;
end;



function TVSComunicacionParaReplicacion01._Cargar_Archivo_En_Servidor_FTP(
  pRutaLocal, pArchivo: String; pModo: String = ''): Boolean;
begin
  Result := False;
  if IdFTP1.Connected then
  begin
    if pModo = 'CLOSEUP' then
    begin
      //IdFTP1.TransferType := IdFTPCommon.TIdFTPTransferType.ftBinary; //Para transferir archivos de texto
      IdFTP1.TransferType := IdFTPCommon.TIdFTPTransferType.ftBinary; //Para transferir archivos comprimidos
      IdFTP1.ChangeDir('vs');
    end
    ELSE
      IdFTP1.TransferType := IdFTPCommon.TIdFTPTransferType.ftASCII; //Para transferir archivos de texto
    try
    begin
      IdFTP1.Put(pRutaLocal, pArchivo, False);
      Result := True;
    end;
    except
      on EError: Exception do
        _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ 'Falló la carga de ' + pRutaLocal + ' en servidor FTP' + ' Error: ' + EError.Message);
    end;
  end
  else
    _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ 'No se pudo cargar archivo ' + pRutaLocal + ' no está conectado al Servidor');
end;

function TVSComunicacionParaReplicacion01._Generar_Lista_Archivos_Remotos: TStringList;
const
  Mascara = '*.*';
begin
  Result      := TStringList.Create;

  if IdFTP1.Connected then
  try
    //IdFTP1.ChangeDir('vs');
    IdFTP1.List(Result, Mascara, False)
  except
    on EError: Exception do
      _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ 'Falló creación de lista de archivos en servidor FTP' + ' Error: ' + EError.Message);
  end;
end;

function TVSComunicacionParaReplicacion01._Get_Pos_log_procesado_Grabado_Destino(
  pSentencias_comunicacion: TStringList;
  pId_archivo_log_procesado_Grabado_Destino: String;
  pPos_log_procesado_Grabado_Destino: Integer; pSentencias_Dep_Texto_Enriquecido: TArrInteger): Longint;
var
  bEsIgual: Boolean;
  sCadena: String;
  iPosIniBegin: Longint;
  slSentenciasLocal: TStringList;
  iPosicion_caracter_inicial: Integer;
  iItem, iItem_Array: Integer;
  sSentencia_Dep_Texto_Enriquecido_Local, sSentencia_Dep_Texto_Enriquecido_Destino: WideString;
  bEsSentencias_Dep_Texto_Enriquecido: Boolean;

  bBuclePrincipal: Boolean;
  //iCantidad_Sentencias_Encontradas, iCantidad_Max_Sentencias_A_Encontrar: Smallint;
  iPosRecorrido: Longint;
  iEnd_log_pos_real: Longint;
  iNro_Extension: Integer;
begin
  Result:= -10;
  iPosIniBegin:= -1;
  //pCadenaDiferencia:= '';

  bBuclePrincipal:= True;
  //iCantidad_Sentencias_Encontradas: 0;
  //Verifico hasta esta cantidad de sentencias, reduce los tiempos y seguir no tiene sentido
  //iCantidad_Max_Sentencias_A_Encontrar:= 10;
  iPosRecorrido:= pPos_log_procesado_Grabado_Destino;
  slSentenciasLocal:= TStringList.Create;
  while bBuclePrincipal=True do
  begin

    //////////////////////

    if qrySelect_Binlog_events_Destino.Active then
      qrySelect_Binlog_events_Destino.Close;
    //qrySelect_Binlog_events_Destino.SQL.Clear;
    //qrySelect_Binlog_events_Destino.sql.Add('SHOW BINLOG EVENTS IN '+  QuotedStr(pId_archivo_log_procesado_Grabado_Destino) + ' FROM '+ IntToStr(iPosRecorrido)+' LIMIT 100');
    //qrySelect_Binlog_events_Destino.sql.Strings[0]:='SHOW BINLOG EVENTS IN '+  QuotedStr(pId_archivo_log_procesado_Grabado_Destino) + ' FROM '+ IntToStr(iPosRecorrido)+' LIMIT 100';

    qrySelect_Binlog_events_Destino.Parameters.ParamByName('pLog_name').Value  := pId_archivo_log_procesado_Grabado_Destino;
    qrySelect_Binlog_events_Destino.Parameters.ParamByName('pPos').Value       := iPosRecorrido;


    try
      // aqui da el problema aveces no ejecuta el open

      //_Grabar_LogErrores_Comunicacion(0, 'Conexion 0  '+pId_archivo_log_procesado_Grabado_Destino+'  '+inttostr(iPosRecorrido));


      //
      // ojo -- Cuando da error 8---Sucursal:101 - Administracion Mitre -  No pudo ejecutar sentencias SHOW BINLOG EVENT en destino. Pos: mysql-bin-vs.000024 26208404 - Pos origen: mysql-bin-vs.000261 297265506
      // Comentar
      //Conexion_Remota.Connected:= False;
      //Conexion_Remota.Connected:= True;
      //

      //_Grabar_LogErrores_Comunicacion(0, 'Conexion 1  '+pId_archivo_log_procesado_Grabado_Destino+'  '+inttostr(iPosRecorrido));

      qrySelect_Binlog_events_Destino.Open;

      //_Grabar_LogErrores_Comunicacion(0, 'Conexion 2  '+pId_archivo_log_procesado_Grabado_Destino+'  '+inttostr(iPosRecorrido));
    except
      on E: Exception do
      begin

        //_Grabar_LogErrores_Comunicacion(0, 'Conexion 3  '+pId_archivo_log_procesado_Grabado_Destino+'  '+inttostr(iPosRecorrido));

        Result:= -2;
        if Assigned(slSentenciasLocal) then
          FreeAndNil(slSentenciasLocal);
        Exit;
      end;
    end;

    //_Grabar_LogErrores_Comunicacion(0, 'Conexion 4  '+pId_archivo_log_procesado_Grabado_Destino+'  '+inttostr(iPosRecorrido));

    with qrySelect_Binlog_events_Destino do
    begin
      if RecordCount > 1 then
      begin

        First;
        while Not Eof do
        begin
          // Nueva forma de sacar el End_log_pos porque en mariadb windows viene en 0 cuando no es el final de la transaccion
          if FieldByName('End_log_pos').AsInteger>0 then
            iEnd_log_pos_real:= FieldByName('End_log_pos').AsInteger
          else
          begin
            if RecNo+1 <= RecordCount then
            begin
              Next;
              iEnd_log_pos_real:= FieldByName('pos').AsInteger;
              Prior;
            end
            else
              iEnd_log_pos_real:= FieldByName('pos').AsInteger;
          end;
          //
          iPosRecorrido:=iEnd_log_pos_real;

          //iPosRecorrido:= FieldByName('end_log_pos').AsInteger;

          //_Grabar_LogErrores_Comunicacion(0, '0  '+pId_archivo_log_procesado_Grabado_Destino+'  '+inttostr(iPosRecorrido));

          //Event_type en el MySql me identifica de que tipo de sentencia se trata si es = Xid se trata de un commit si es = query se trata de un insert o update o delete
          if ((LeftStr(FieldByName('info').AsString, 6 ) = 'COMMIT') and ( (LowerCase(FieldByName('Event_type').AsString) = 'xid') or (LowerCase(FieldByName('Event_type').AsString) = 'query') or (LowerCase(FieldByName('Event_type').AsString) = 'annotate_rows') )) then
          begin
            // Aqui pongo en -1 para saber si encontro sentencia en el registro de eventos de la BD destino
            Result:= -1;
            // Verifico si es el mismo conjunto de sentencia guardado en el destino
            if (slSentenciasLocal.Count-2) = pSentencias_comunicacion.Count then
            begin
              bEsIgual:= True;
              //Hasta Count-3 porque no debe tomar en cuenta las 2 ultimas sentencias, que es la insersion en log_comunicacion_externo
              //y la actualizacion en sucursales del ultimo pos procesado origen  (pos_log_procesado_origen y id_archivo_log_procesado_origen)

              for iItem := 0 to slSentenciasLocal.Count-3 do
              begin
                // ojo
                // Controla hasta 4 sentencias
                //if iItem > 3 then
                //  Break;


                bEsSentencias_Dep_Texto_Enriquecido:= False;

                for iItem_Array := Low(pSentencias_Dep_Texto_Enriquecido) to High(pSentencias_Dep_Texto_Enriquecido) do
                begin
                  if pSentencias_Dep_Texto_Enriquecido[iItem_Array] = iItem then
                  begin
                    bEsSentencias_Dep_Texto_Enriquecido:= True;
                  end;
                end;

                if bEsSentencias_Dep_Texto_Enriquecido = True then
                begin
                  sSentencia_Dep_Texto_Enriquecido_Local:= _Depurar_Texto_Erriquecido(slSentenciasLocal.Strings[iItem]);
                  sSentencia_Dep_Texto_Enriquecido_Destino:= _Depurar_Texto_Erriquecido(pSentencias_comunicacion.Strings[iItem]);
                  //if sSentencia_Dep_Texto_Enriquecido_Local<> sSentencia_Dep_Texto_Enriquecido_Destino then
                  if CorrijeCadenaParticular(LeftStr(sSentencia_Dep_Texto_Enriquecido_Local,100))<> CorrijeCadenaParticular(LeftStr(sSentencia_Dep_Texto_Enriquecido_Destino,100)) then
                  begin
                    bEsIgual:= False;
                    Break;
                  end;
                end
                else
                begin
                  if CorrijeCadenaParticular(Trim(slSentenciasLocal.Strings[iItem])) <> CorrijeCadenaParticular(Trim(pSentencias_comunicacion.Strings[iItem])) then
                  begin
                    bEsIgual:= False;
                    //pCadenaDiferencia:= pCadenaDiferencia + 'Item Sentencia:' +IntToStr(iItem)+ Chr(13)+
                    //  'Sentencia origen:'+Chr(13)+CorrijeCadenaParticular(Trim(pSentencias_comunicacion.Strings[iItem]))+Chr(13)+Chr(13)+'Sentencia destino:'+Chr(13)+CorrijeCadenaParticular(Trim(slSentenciasLocal.Strings[iItem]));
                    Break;
                  end;
                end;
              end;
              if bEsIgual then  // Encontro el conjunto de sentencias
              begin
                Result:= iPosIniBegin;
                if Assigned(slSentenciasLocal) then
                  FreeAndNil(slSentenciasLocal);
                Exit;
              end;
            end
            else
            begin
              // Nada: sigue con Result:= -1;
              //Result:= -3;
              Result:= -1;
            end;
            //

          end
          else
          begin
            //if LowerCase(FieldByName('Event_type').AsString) =  'query' then
            if ((LowerCase(FieldByName('Event_type').AsString)='query') or (LowerCase(FieldByName('Event_type').AsString) = 'annotate_rows') or (LowerCase(FieldByName('Event_type').AsString)='gtid')) then
            begin
              sCadena:= FieldByName('info').AsString;

              //*********Identifico si se trata de un insert o delete o update para guardar la cadena en el sql. Tambien un Begin para indicar iPosIniBegin**********//
              //PosEx me devuelve la posición del primer caracter de una subcadena especificada
              if ((PosEx('insert', LowerCase(sCadena))> 0) and (PosEx('into', LowerCase(sCadena))> 0 ))then
              begin
                sCadena:= RightStr(sCadena,(Length(sCadena) - (PosEx('insert', LowerCase(sCadena))-1)) );
                iPosicion_caracter_inicial:=PosEx('into', LowerCase(sCadena));
                if iPosicion_caracter_inicial> 0  then
                begin
                  slSentenciasLocal.Add(sCadena);
                end;//end if iPosicion_caracter_inicial> 0  then
              end //end if  PosEx('insert', sCadena)> 0  then
              else if ((PosEx('delete', LowerCase(sCadena))>0) and (PosEx('from', LowerCase(sCadena))>0)) then
              begin
                sCadena:= RightStr(sCadena,(Length(sCadena) - (PosEx('delete', LowerCase(sCadena))-1)) );
                iPosicion_caracter_inicial:=PosEx('from', LowerCase(sCadena));
                if iPosicion_caracter_inicial> 0  then
                begin
                  slSentenciasLocal.Add(sCadena);
                end;//end if iPosicion_caracter_inicial> 0  then
              end //end if  PosEx('delete', sCadena)> 0  then
              else if ((PosEx('update', LowerCase(sCadena))>0) and (PosEx('set', LowerCase(sCadena))>0)) then
              begin
                iPosicion_caracter_inicial:=PosEx('update', LowerCase(sCadena));
                if  iPosicion_caracter_inicial> 0  then
                begin
                  sCadena:= RightStr(sCadena,(Length(sCadena) - (iPosicion_caracter_inicial-1)) );
                  if PosEx('set', LowerCase(sCadena))> 0  then
                  begin
                    slSentenciasLocal.Add(sCadena);
                  end;//end if if PosEx('set', sCadena)> 0  then
                end; //end if iPosicion_caracter_inicial> 0  then
              end //if PosEx('update', LowerCase(sCadena))> 0  then
              else if PosEx('begin', LowerCase(sCadena))> 0  then
              //else if LeftStr(sCadena, 5 ) = 'BEGIN' then
              begin

                slSentenciasLocal.Clear;
                iPosIniBegin:= FieldByName('pos').AsInteger;
              end; //if PosEx('update', LowerCase(sCadena))> 0  then
              //*********Identifico si se trata de un insert o delete o update para guardar la cadena en el sql. Tambien un Begin para indicar iPosIni**********//

            end
            else
            begin
              if LowerCase(FieldByName('Event_type').AsString) =  'stop' then
              begin
                //bFin_archivo:= True;
              end;
            end; //end if   LowerCase(FieldByName('Event_type').AsString) =  'query' then
          end;//end if  LeftStr(FieldByName('info').AsString, 6 ) = 'COMMIT' then
          Next;
        end; //while Not Eof do

      end
      else
      begin
        if _Existe_log_Posterior_Destino(pId_archivo_log_procesado_Grabado_Destino) then
        begin
          iNro_Extension:= StrToInt(RightStr(pId_archivo_log_procesado_Grabado_Destino, oParametro.Comu_largo_extension_file_binlog)) + 1;

          if Self.Tag<>100 then
            pId_archivo_log_procesado_Grabado_Destino:= oParametro.Comu_nombre_file_binlog + '.'+ RightStr('00000000000000000'+IntToStr(iNro_Extension), oParametro.Comu_largo_extension_file_binlog)
          else
            // En AWS: mysql-bin-changelog - No se puede configurar
            pId_archivo_log_procesado_Grabado_Destino:= 'mysql-bin-changelog' + '.'+ RightStr('00000000000000000'+IntToStr(iNro_Extension), oParametro.Comu_largo_extension_file_binlog);

          pPos_log_procesado_Grabado_Destino:=0;
          iPosRecorrido:= pPos_log_procesado_Grabado_Destino;

          //_Grabar_LogErrores_Comunicacion(0, '2  '+pId_archivo_log_procesado_Grabado_Destino+'  '+inttostr(iPosRecorrido));

        end
        else
        begin
          // Va con -1 porque no se pudo verificar

          //_Grabar_LogErrores_Comunicacion(0, '3  '+pId_archivo_log_procesado_Grabado_Destino+'  '+inttostr(iPosRecorrido));

          if Assigned(slSentenciasLocal) then
            FreeAndNil(slSentenciasLocal);
          Exit;
        end;
      end;  //if RecordCount > 0 then
    end;  //with qrySelect_Binlog_events_Destino do

    //_Grabar_LogErrores_Comunicacion(0, 'nn  Llego al bucle '+pId_archivo_log_procesado_Grabado_Destino+'  '+inttostr(iPosRecorrido));

  end;  //while bBuclePrincipal=True do
  if Assigned(slSentenciasLocal) then
    FreeAndNil(slSentenciasLocal);
end;


procedure TVSComunicacionParaReplicacion01._Hab_des_Timer(pTag: Integer;
  bValor: Boolean);
begin
  if pTag= 1 then
    Timer1.Enabled:= bValor
  else if pTag= 2 then
    Timer2.Enabled:= bValor
  else if pTag= 3 then
    Timer3.Enabled:= bValor
  else if pTag= 4 then
    Timer4.Enabled:= bValor
  else if pTag= 5 then
    Timer5.Enabled:= bValor
  else if pTag= 6 then
    Timer6.Enabled:= bValor
  else if pTag= 7 then
    Timer7.Enabled:= bValor
  else if pTag= 8 then
    Timer8.Enabled:= bValor
  else if pTag= 9 then
    Timer9.Enabled:= bValor
  else if pTag= 10 then
    Timer10.Enabled:= bValor
  else if pTag= 11 then
    Timer11.Enabled:= bValor
  else if pTag= 12 then
    Timer12.Enabled:= bValor
  else if pTag= 13 then
    Timer13.Enabled:= bValor
  else if pTag= 14 then
    Timer14.Enabled:= bValor
  else if pTag= 15 then
    Timer15.Enabled:= bValor
  else if pTag= 16 then
    Timer16.Enabled:= bValor
  else if pTag= 17 then
    Timer17.Enabled:= bValor
  else if pTag= 18 then
    Timer18.Enabled:= bValor
  else if pTag= 19 then
    Timer19.Enabled:= bValor
  else if pTag= 20 then
    Timer20.Enabled:= bValor
  else if pTag= 21 then
    Timer21.Enabled:= bValor
  else if pTag= 22 then
    Timer22.Enabled:= bValor
  else if pTag= 23 then
    Timer23.Enabled:= bValor
  else if pTag= 24 then
    Timer24.Enabled:= bValor
  else if pTag= 25 then
    Timer25.Enabled:= bValor
  else if pTag= 26 then
    Timer26.Enabled:= bValor
  else if pTag= 27 then
    Timer27.Enabled:= bValor
  else if pTag= 28 then
    Timer28.Enabled:= bValor
  else if pTag= 100 then
    Timer100.Enabled:= bValor;
end;

procedure TVSComunicacionParaReplicacion01._Informar_IQVIA_externo(
  pId_Empresa_Archivo: Integer; pId_Sucursal_Archivo: Integer; pUbicacion_archivos_origen,
  pUbicacion_archivos_procesados: String);
var
  oIQVIA                          : TIQVIA;
  sRespuesta                      : String;
  sRespuesta_Query                : String;
  oSucursal                       : TSucursal;
  dtFecha                         : TDate;
  oGestor_Tipo_Validacion_Receta: TGestor_Tipo_Validacion_Receta;
  oValidadores_Por_Sucursal:TValidadores_Por_Sucursal;
  oTipo_Validacion_Receta  : TTipo_Validacion_Receta;
begin
  _Grabar_LogErrores_Comunicacion(-2, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Inicializando proceso... ');

  oSucursal:= _Gestor_Sucursal._Buscar(pId_Empresa_Archivo, pId_Sucursal_Archivo);

  oGestor_Tipo_Validacion_Receta:= TGestor_Tipo_Validacion_Receta.Create;
  oTipo_Validacion_Receta       := oGestor_Tipo_Validacion_Receta._Buscar(1); // Imed

  if Assigned(oGestor_Tipo_Validacion_Receta) then
    FreeAndNil(oGestor_Tipo_Validacion_Receta);

  if Assigned(oTipo_Validacion_Receta) then
  begin
    if Assigned(oSucursal) then
    begin
      oValidadores_Por_Sucursal:= oTipo_Validacion_Receta._Get_Validadores_Por_Sucursal(oSucursal.Id_empresa, oSucursal.Id_sucursal);
      if Assigned(oValidadores_Por_Sucursal) then
      begin
        _Grabar_LogErrores_Comunicacion(-2, 'Sucursal: '+ IntToStr(pId_Sucursal_Archivo)+' - '+oSucursal.Des_sucursal+' - Carpeta origen:'+pUbicacion_archivos_origen+' - Carpeta procesados:'+pUbicacion_archivos_procesados);

        oIQVIA     := TIQVIA.Create(Nil, oParametro, oValidadores_Por_Sucursal.Adesfa_Codigo_Prestador, pUbicacion_archivos_procesados+'\Errores', pUbicacion_archivos_origen, pUbicacion_archivos_procesados);

        dtFecha:= Now;
        try
          sRespuesta := oIQVIA._GenerarSolicitud_IQVIA(FormatDateTime('yyyymmdd', dtFecha), oSucursal.Id_empresa, oSucursal.Id_sucursal, False, 'WS');
        except
          on E: Exception do
          begin
            sRespuesta := 'Error al procesar "_GenerarSolicitud_IQVIA" -  Error: '+ (E.Message);
          end;
        end;

        if Assigned(oIQVIA) then
          FreeAndNil(oIQVIA);

        if Length(Trim(sRespuesta))>0  then
        begin
          _Grabar_LogErrores_Comunicacion(-2, sRespuesta);
        end;

        // Libero oValidadores_Por_Sucursal
        if Assigned(oValidadores_Por_Sucursal) then
          FreeAndNil(oValidadores_Por_Sucursal);
      end;
    end;
    _Grabar_LogErrores_Comunicacion(-2, 'Proceso Finalizado!');
  end;

  if Assigned(oTipo_Validacion_Receta) then
    FreeAndNil(oTipo_Validacion_Receta);
end;


function TVSComunicacionParaReplicacion01._Modificar_Comu_Fecha_Hora(
  pId_Sucursal: Integer): String;
var
  dFecha_Hora : TDateTime;
  dtFechaHoraUltimaRegistroLog: TDateTime;
begin
  Result := '';
  try

    // Saco la Hora actual de la BD
    if qrySelect_Now.Active then
      qrySelect_Now.Close;
    qrySelect_Now.Open;
    if qrySelect_Now.RecordCount > 0 then
      dFecha_Hora:= qrySelect_Now.FieldByName('Fecha_Hora').AsDateTime
    else
      dFecha_Hora:= Now;
    qrySelect_Now.Close;
    //

    // Consulto Fecha_hora ultima actualizacion
    if qrySelect_Comu_Fecha_Hora.Active then
      qrySelect_Comu_Fecha_Hora.Close;
    qrySelect_Comu_Fecha_Hora.Parameters.ParamByName('pId_sucursal').Value:= pId_Sucursal;
    qrySelect_Comu_Fecha_Hora.Open;

    if qrySelect_Comu_Fecha_Hora.RecordCount > 0 then
      dtFechaHoraUltimaRegistroLog:= qrySelect_Comu_Fecha_Hora.FieldByName('Fecha_Hora').AsDateTime
    else
      dtFechaHoraUltimaRegistroLog:= dFecha_Hora - 1;  // Para que exista diferencia y grabe
    qrySelect_Comu_Fecha_Hora.Close;
    //

    if SecondsBetween(dFecha_Hora, dtFechaHoraUltimaRegistroLog) > 90 then
    begin
      qryUpdate_Comu_Fecha_Hora.Parameters.ParamByName('pId_sucursal').Value := pId_sucursal;
      qryUpdate_Comu_Fecha_Hora.Parameters.ParamByName('pFecha_hora').Value := FormatDateTime('yyyy/mm/dd HH:mm:ss', dFecha_Hora);
      qryUpdate_Comu_Fecha_Hora.Connection.BeginTrans;
      qryUpdate_Comu_Fecha_Hora.ExecSQL;
      qryUpdate_Comu_Fecha_Hora.Connection.CommitTrans;

      dtFechaHoraUltimaRegistroLog:= dFecha_Hora;
    end;
  except
    on E: Exception do
    begin
      Result := 'Error al actualizar Fecha_Hora en sucursales_comunicacion -  Error: '+ E.Message;
      if qryUpdate_Comu_Fecha_Hora.Connection.InTransaction then
        qryUpdate_Comu_Fecha_Hora.Connection.RollbackTrans;
    end;
  end;

end;

function TVSComunicacionParaReplicacion01._Modificar_Comu_Fecha_Hora_Ultima_Revision(pId_Sucursal: Integer): String;
begin
  with qryUpdate_Comu_Fecha_Hora_Ultima_Revision do
  begin
    Result := '';
    try
      Parameters.ParamByName('pId_sucursal').Value := pId_sucursal;
      Connection.BeginTrans;
      ExecSQL;
      Connection.CommitTrans;
    except
      Result := 'Error al actualizar Comu_fecha_hora_ultima_revision en sucursales_comunicacion';
      Connection.RollbackTrans;
    end;
  end;
end;

function TVSComunicacionParaReplicacion01._Existe_Mas_Sentencias_del_mismo_log(
  pId_archivo_log_procesado: String; pUltimo_Pos : Longint): Boolean;
begin
  with qrySelect_Binlog_events do
  begin
    sql.Clear;
    sql.Add('SHOW BINLOG EVENTS IN '+  QuotedStr(pId_archivo_log_procesado) + 'FROM '+ IntToStr(pUltimo_Pos) +' LIMIT 10');
    //sql.Add('SELECT * FROM bin_log_events_tabla WHERE bin_log_events_tabla.Log_name= '+ QuotedStr(pId_archivo_log_procesado) +' LIMIT 10');
    try
      Open;
      if RecordCount > 0 then
        Result:= True
      else
        Result:= False;
    except
      Result:= False;
    end;
    close;
  end;//end with qrySelect_Binlog_events do
end;

procedure TVSComunicacionParaReplicacion01._Exportacion_DataView;
var
  dtFechaActual : TDateTime;
  dtFechaDesde,
  dtFechaHasta,
  dtFechaEnviar  : TDate;
  wDia, wMes, wAnio, wHora, wMin, wSeg, wMil : Word;
  sMensaje            : String;
  oLog_Informe_Auto   : TLog_Informe_Automatico;
  oSucursal : TSucursal;
  iTotal : Integer;
  oHttpClient: THttpClient;
  Response: IHTTPResponse;
  sUrl, sJsonBody: string;
  ssBodyVacio: TStringStream;
  Headers: TNetHeaders;
  sToken, sFecha_Caducidad, sRespuesta_Credenciales: String;
  sFrom, sTo, sValue: String;
  dsTodasSucursales : TDataSet;
  oDatasetToExcel : TDataSetToExcel;
begin
  // Deshabilito
  TimerExportacionDataview.Enabled:= False;

  // Saco la Hora actual de la BD
  if qrySelect_Now.Active then
    qrySelect_Now.Close;
  qrySelect_Now.Open;
  if qrySelect_Now.RecordCount > 0 then
    dtFechaActual:= qrySelect_Now.FieldByName('Fecha_Hora').AsDateTime
  else
    dtFechaActual:= Now;
  qrySelect_Now.Close;

  DecodeDateTime(dtFechaActual, wAnio, wMes, wDia, wHora, wMin, wSeg, wMil);

  if wHora > 0 then  //Espero que cambie el día para exportar lo del día que se completó, debo subir lo del día de ayer
  begin
    dtFechaHasta := EncodeDate(wAnio, wMes, wDia); //Hoy
    dtFechaDesde := dtFechaHasta-2; //Anteayer
  end;

  dsTodasSucursales := nil;
  with qryTodas_Sucursales do
  begin
    if Active then Close;
    Open;
    if RecordCount > 0 then
      dsTodasSucursales := qryTodas_Sucursales;
  end;

  if Assigned(dsTodasSucursales) then
  begin
    if dsTodasSucursales.RecordCount > 0 then
    begin

//      oDatasetToExcel := TDataSetToExcel.Create(dsTodasSucursales,'C:\Rta\sucursales.xls');
//      oDatasetToExcel.ExportarXLS(True);
      oDatasetToExcel.Free;

      dsTodasSucursales.First;
      while not dsTodasSucursales.Eof do
      begin

        if _Gestor_Log_Informe_Automatico._Periodo_Cargado(dsTodasSucursales.FieldByName('id_empresa').AsInteger, dsTodasSucursales.FieldByName('id_sucursal').AsInteger, 'DVW', dtFechaDesde, dtFechaHasta) then
        begin
          dsTodasSucursales.Next;
          Continue;
        end;

        with qrySelectTicketsDia do
        begin
          Parameters.ParamByName('pId_Empresa').Value  := dsTodasSucursales.FieldByName('id_empresa').AsInteger;
          Parameters.ParamByName('pId_Sucursal').Value := dsTodasSucursales.FieldByName('id_sucursal').AsInteger;
          Parameters.ParamByName('pdechaDesde').Value  := FormatDateTime('yyyy/mm/dd',dtFechaDesde);
          Parameters.ParamByName('pFechaHasta').Value  := FormatDateTime('yyyy/mm/dd',dtFechaHasta);

          if Active then Close;
          Open;
      //
          if qrySelectTicketsDia.RecordCount > 0  then
          begin
            iTotal :=  qrySelectTicketsDia.FieldByName('cantidad_comp').AsInteger;

            oSucursal := nil;

            if dsTodasSucursales.FieldByName('id_sucursal').AsInteger = 7 then
              Sleep(1000);

            oSucursal := TSucursal.Create(dsTodasSucursales.Fields);

            if Assigned(oSucursal) then
            begin
              if not ((Length(oSucursal.DV_User) > 0)  and (Length(oSucursal.DV_Pass) > 0) and (Length(oSucursal.DV_Id_Sucursal) > 0)) then
              begin
                dsTodasSucursales.Next;
                Continue;
              end;

              if Length(oSucursal.DV_Access_Token) > 0 then
                sToken := oSucursal.DV_Access_Token
              else
              begin
                try
                  oHttpClient := THttpClient.Create;
                  sUrl := oParametro.DV_URL_Base + '/api/v1/auth/login/?email=' + oSucursal.DV_User + '&password=' + oSucursal.DV_Pass;

                  // Crear un contenido vacío para el cuerpo de la solicitud, que no usa un body
                  ssBodyVacio := TStringStream.Create('', TEncoding.UTF8);
                  try
                    Response := oHttpClient.Post(sUrl, ssBodyVacio, nil, [TNameValuePair.Create('Content-Type', 'application/json')]);

                    if Response.StatusCode = 200 then
                    begin
                      sToken := DevuelveValorDeClaveEnJSON( Response.ContentAsString(TEncoding.UTF8), 'token');
                      //Para mas adelante cuando se pueda usar refresh token, adaptar esta parte
                      //para guardar ese dato y la caducidad real, no esta que es inventada
                      //porque no se vence el access token pero ya estan los campos para guardarlos,
                      //solamente falta que se parseen los valores cuando lleguen, para gaurdarlos
                      sFecha_Caducidad := FormatDateTime('yyyy-mm-dd hh:mm:ss', dtFechaActual+180);
                      sRespuesta_Credenciales := _Gestor_Sucursal._Actualizar_Credenciales_DataView(oSucursal.Id_Empresa, oSucursal.Id_Sucursal, sToken, '', sFecha_Caducidad);
                    end
                    else
                      _Grabar_LogErrores_Comunicacion(-10, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - ' + 'Error: ' + 'Status: ' + IntToStr(Response.StatusCode) + ' - Mensaje: ' + Response.StatusText + ' Para sucursal ' + IntToStr(oSucursal.Id_sucursal));
                  finally
                    oHttpClient.Free;
                  end;
                finally
                end;
              end;

              if Length(sToken) > 0 then
              begin
                try
                  sUrl := oParametro.DV_URL_Base + '/api/v1/variable/'+ oSucursal.DV_Id_Sucursal;

                  dtFechaEnviar := (EncodeDateTime(wAnio, wMes, wDia, 0 , 0, 0, 0)-1);
                  sFrom := FormatDateTime('yyyy-mm-dd hh:mm:ss', dtFechaEnviar) ;

                  dtFechaEnviar := (EncodeDateTime(wAnio, wMes, wDia, 23 , 59, 59, 0)-1);
                  sTo := FormatDateTime('yyyy-mm-dd hh:mm:ss', dtFechaEnviar) ;

                  sJsonBody := '[{"from":"' + sFrom + '", "to":"' + sTo + '", "value":' + IntToStr(iTotal) + '}]';

                  oHttpClient := THttpClient.Create;

                  SetLength(Headers, 3);
                  Headers[0].Name := 'Content-Type';
                  Headers[0].Value := 'application/json';
                  Headers[1].Name := 'Authorization';
                  Headers[1].Value := 'Bearer ' + sToken;
                  Headers[2].Name := 'Accept';
                  Headers[2].Value := 'application/json';

                  Response := oHttpClient.Put(sUrl, TStringStream.Create(sJsonBody, TEncoding.UTF8), nil, Headers);

                  if Response.StatusCode = 200 then
                  begin
                    oLog_Informe_Auto := TLog_Informe_Automatico.Create('DVW', dtFechaDesde, dtFechaHasta, Response.ContentAsString(TEncoding.UTF8) + ' | '+ 'Total Comprobantes: ' + IntToStr(iTotal), oSucursal.Id_Empresa, oSucursal.Id_Sucursal);
                    sMensaje := _Gestor_Log_Informe_Automatico._Insertar(oLog_Informe_Auto);
                    if sMensaje = '' then
                      _Grabar_LogErrores_Comunicacion(-10, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Se cargó la información de sucursal ' + IntToStr(oSucursal.Id_sucursal) + ' para el día ' + FormatDateTime('yyyy-mm-dd', Now-1) + ' Total Comprobantes: ' + IntToStr(iTotal))
                    else
                      _Grabar_LogErrores_Comunicacion(-10, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - ' + sMensaje + ' En sucursal ' + IntToStr(oSucursal.Id_sucursal));
                    if Assigned(oLog_Informe_Auto) then
                     FreeAndNil(oLog_Informe_Auto);
                  end
                  else
                    _Grabar_LogErrores_Comunicacion(-10, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - ' + 'Error, no se pudo cargar datos a DataView para sucursal' + IntToStr(oSucursal.Id_sucursal)  + ' por: ' + Response.StatusText);
                finally
                  oHttpClient.Free;
                end;

              end
              else
                _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Falló al intentar obtener el Access Token, tanto localmente como vía API para Sucursal: ' + IntToStr(oSucursal.Id_sucursal));


            end
            else
                _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Falló al instanciar la Sucursal '  + IntToStr(oSucursal.Id_sucursal) + ' para obtener parámetros');
          end
          else
            _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - No existen datos para exportar en sucursal' + IntToStr(osucursal.Id_sucursal));

          Close;
        end;
        dsTodasSucursales.Next;
      end;

    end;
  end;

  // Habilito
  TimerExportacionDataview.Enabled:= True;

end;

procedure TVSComunicacionParaReplicacion01._Exportar_Datos_Receta;
const
  RutaArchCloseUp = 'c:\vs-comunicacion\CloseUp\';
var
  _InterfaceArchivos  : TInterfaceArchivos;

  bExporto,
  bCargado,
  bExisteLocal,
  bExisteRemoto       : Boolean;
  oEmpresa            : TEmpresa;
  dtFechaActual       : TDateTime;
  dtFechaDesde,
  dtFechaHasta,
  dtFecha_Inicio_CUp  : TDate;
  sNombreArchivo,
  sNombre_Archivo_Zip  : String;
  wDia, wMes, wAnio,
  wHora, wMin, wSeg,
  wMil                : Word;
  slArchivosRemotos,
  slTextoCompleto,
  slTextoFiltrado     : TStringList;
  iLinea              : Integer;
  sMensaje            : String;
  oLog_Informe_Auto   : TLog_Informe_Automatico;
  oFormato_Fecha      : TFormatSettings;
  ArchivoZip          : TZipFile;
begin
  // Deshabilito
  TimerExportarRecetasCloseup.Enabled:= False;

  dtFecha_Inicio_CUp := _Gestor_Log_Informe_Automatico._Fecha_Inicio_CloseUp;
  if not(dtFecha_Inicio_CUp > 0) then
  begin
    TimerExportarRecetasCloseup.Enabled:= True;
    Exit;
  end;

  // Saco la Hora actual de la BD
  if qrySelect_Now.Active then
    qrySelect_Now.Close;
  qrySelect_Now.Open;
  if qrySelect_Now.RecordCount > 0 then
    dtFechaActual:= qrySelect_Now.FieldByName('Fecha_Hora').AsDateTime
  else
    dtFechaActual:= Now;
  qrySelect_Now.Close;
  //

  if TDate(dtFechaActual) < TDate(dtFecha_Inicio_CUp) then
  begin
    TimerExportarRecetasCloseup.Enabled:= True;
    Exit;
  end;

  DecodeDateTime(dtFechaActual, wAnio, wMes, wDia, wHora, wMin, wSeg, wMil);

  if wHora > 17 then  //Para las 17 hs. ya se habrán cerrado todos los tickets del día anterior
  begin
    if (wDia >= 5) and (wDia <= 11) then  // Cierra semana del 28 al 4
    begin
      if wMes = 1 then //me fijo si cambió el año
      begin
        dtFechaDesde := EncodeDate(wAnio-1, 12, 28);
        dtFechaHasta := EncodeDate(wAnio, wMes, 4);
      end
      else
      begin
        dtFechaDesde := EncodeDate(wAnio, wMes-1, 28);
        dtFechaHasta := EncodeDate(wAnio, wMes, 4);
      end;
    end
    else if (wDia >= 12) and (wDia <= 18) then  // Cierra semana del 5 al 11
    begin
      dtFechaDesde := EncodeDate(wAnio, wMes, 5);
      dtFechaHasta := EncodeDate(wAnio, wMes, 11);
    end
    else if (wDia >= 19) and (wDia <= 27) then  // Cierra semana del 12 al 18
    begin
      dtFechaDesde := EncodeDate(wAnio, wMes, 12);
      dtFechaHasta := EncodeDate(wAnio, wMes, 18);
    end
    else if (wDia >= 28) or (wDia <= 4) then // Cierra semana del 19 al 27
    begin
      if (wDia >= 1) and (wDia <= 4) then
      begin
        dtFechaDesde := EncodeDate(wAnio, wMes-1, 19);
        dtFechaHasta := EncodeDate(wAnio, wMes-1, 27);
      end
      else
      begin
        dtFechaDesde := EncodeDate(wAnio, wMes, 19);
        dtFechaHasta := EncodeDate(wAnio, wMes, 27);
      end;
    end
    else // cualquier otro dia que no requiera exportar recetas
    begin
      TimerExportarRecetasCloseup.Enabled:= True;
      Exit;
    end;

    oEmpresa := _Gestor_Empresa._Buscar(oParametro.Id_Empresa);
    oEmpresa.Des_empresa := CorrijeCadena(oempresa.Des_empresa);
    oEmpresa.Des_empresa := UpperCase(ReplaceStr(oEmpresa.Des_empresa, ' ', '_'));

    sNombreArchivo    := 'Recetas_'+CorrijeCadena(oEmpresa.Des_empresa)+'_'+FormatDateTime('yyyymmdd',dtFechaDesde)+'-'+FormatDateTime('yyyymmdd',dtFechaHasta)+'.txt';

    sNombre_Archivo_Zip:= ReplaceStr(sNombreArchivo,'.txt','');
    sNombre_Archivo_Zip:= sNombre_Archivo_Zip + '.zip';

    if _Gestor_Log_Informe_Automatico._Recetas_Cargadas(sNombre_Archivo_Zip) then
    begin
      TimerExportarRecetasCloseup.Enabled:= True;
      Exit;
    end;

    bExporto := False;
    bCargado := False;

    with qrySelectExportacion do
    begin
      //Parameters.ParamByName('pId_Empresa').Value := oParametro.Id_Empresa;
      Parameters.ParamByName('pFecha_comprobante_ini').Value := FormatDateTime('yyyy/mm/dd',dtFechaDesde);
      Parameters.ParamByName('pFecha_comprobante_fin').Value := FormatDateTime('yyyy/mm/dd',dtFechaHasta);

      if Active then
        Close;
      Open;

      if qrySelectExportacion.RecordCount > 0  then
      begin
        _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Se generó el conjunto de datos con un total de ' + IntToStr(qrySelectExportacion.RecordCount)  + ' recetas' );
        if not(DirectoryExists(RutaArchCloseUp)) then
        begin
          try
            CreateDir(RutaArchCloseUp)
          except
            begin
              _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - No se pudo crear carpeta ' + RutaArchCloseUp);
              TimerExportarRecetasCloseup.Enabled:= True;
              Exit;
            end;
          end;
        end;

        _InterfaceArchivos:= TInterfaceArchivos.Create(oParametros_Formato.Formato_exportacion_recetas);
        bExporto := _InterfaceArchivos.Exportar(Nil, qrySelectExportacion,Nil,'', RutaArchCloseUp+sNombreArchivo, True);

        if Assigned(_InterfaceArchivos) then
          FreeAndNil(_InterfaceArchivos);

        if bExporto then
        begin
          //Tratamiento para Farmanor
          slTextoCompleto := TStringList.Create;
          slTextoFiltrado := TStringList.Create;
          slTextoCompleto.LoadFromFile(RutaArchCloseUp+sNombreArchivo);
          if slTextoCompleto.Count > 0 then
          begin
            for iLinea := 0 to slTextoCompleto.Count-1 do
            begin
              if ContainsText(slTextoCompleto[iLinea], 'Descuento Farmanor') then
                Continue
              else
                slTextoFiltrado.Add(slTextoCompleto[iLinea]);
            end;
         end;

          slTextoFiltrado.SaveToFile(RutaArchCloseUp+sNombreArchivo);
          FreeAndNil(slTextoCompleto);
          FreeAndNil(slTextoFiltrado);
          _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - El archivo ' + RutaArchCloseUp+sNombreArchivo +' Se generó localmente con éxito');
          //fin de tratamiento Farmanor

          //Conecto al FTP y cargo el archivo
          with oParametro do
            _Conectar_Servidor_FTP(CloseUp_FTP_Host, CloseUp_FTP_User, CloseUp_FTP_Pass, CloseUp_FTP_Port, 3000, True);

          if not (IdFTP1.Connected) then
          begin
            _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - El archivo ' + RutaArchCloseUp+sNombre_Archivo_Zip +' No se pudo cargar, falló la conexión con el servidor');
            _Desconectar_Servidor_FTP;
            TimerExportarRecetasCloseup.Enabled:= True;
            Exit;
          end;

          ArchivoZip := TZipFile.Create;
          try
            try
              ArchivoZip.Open(RutaArchCloseUp+sNombre_Archivo_Zip, zmWrite);
              ArchivoZip.Add(RutaArchCloseUp+sNombreArchivo);

            except
              //Mas abajo resuelvo si existe localmente
            end;
          finally
            FreeAndNil(ArchivoZip)
          end;

          bExisteLocal  := False;
          if FileExists(RutaArchCloseUp+sNombre_Archivo_Zip) then
            bExisteLocal := True;

          if bExisteLocal then
            bCargado := _Cargar_Archivo_En_Servidor_FTP((RutaArchCloseUp+sNombre_Archivo_Zip), sNombre_Archivo_Zip, 'CLOSEUP');

          if bCargado then
          begin
            bExisteRemoto := False;
            slArchivosRemotos := _Generar_Lista_Archivos_Remotos;
            if (Assigned(slArchivosRemotos)) and (slArchivosRemotos.Count > 0) then
              bExisteRemoto     := _Archivo_Cargado_FTP(sNombre_Archivo_Zip, slArchivosRemotos);

            if bExisteLocal and bExisteRemoto then
            begin
              //Inserto en la tabla de hechos registro para archivo nuevo
              oLog_Informe_Auto := TLog_Informe_Automatico.Create('CUP', dtFechaDesde, dtFechaHasta, sNombre_Archivo_Zip, oParametro.Id_Empresa);
              sMensaje := _Gestor_Log_Informe_Automatico._Insertar(oLog_Informe_Auto);
              if sMensaje = '' then
                _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - El archivo ' + RutaArchCloseUp+sNombre_Archivo_Zip +' se cargó en el servidor FTP')
              else
                _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - ' + sMensaje);
              if Assigned(oLog_Informe_Auto) then
               FreeAndNil(oLog_Informe_Auto);
            end
            else
            if bExisteLocal then
              _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Falló al cargar el archivo ' + RutaArchCloseUp+sNombre_Archivo_Zip +' en el servidor');
          end
          else
          begin
            _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Falló al cargar el archivo ' + RutaArchCloseUp+sNombre_Archivo_Zip +' en el servidor');
          end;

        end
        else
            _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - Error al exportar');
      end
      else
        _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ ' - No existen datos para exportar');

      Close;
    end;
  end;

  _Desconectar_Servidor_FTP;

  // Habilito
  TimerExportarRecetasCloseup.Enabled:= True;

end;

function TVSComunicacionParaReplicacion01._Extraer_NombreTabla(
  pCadena: String; pPosicion_caracter_final: Integer): String;
begin
  Result:= '';
  //guardo el nombre de la tabla
  while pPosicion_caracter_final<=(Length(pCadena)) do //mientras que no se llegue a la posición final de la cadena
  begin

    if ((pCadena[pPosicion_caracter_final]=' ') or (pCadena[pPosicion_caracter_final]=#13) or (pCadena[pPosicion_caracter_final]=#10) or (pCadena[pPosicion_caracter_final]='(')) then
    begin
      //Si ya encontre el nombre de la tabla salgo del bucle
      if Length(Result) > 0 then
      begin
        Result:= StringReplace(Result,'`','', [rfReplaceAll]);
        Break;
      end;
    end
    else
    begin
      Result:= Result + pCadena[pPosicion_caracter_final];
    end;
    pPosicion_caracter_final := pPosicion_caracter_final + 1;
  end; //end while pPosicion_caracter_final<=(Length(pCadena)) do //mientras que no se llegue a la posición final de la cadena
end;

function TVSComunicacionParaReplicacion01._Depurar_Texto_Erriquecido(
  pTexto: WideString): WideString;
var
  iItemCaracter: Integer;
  iCantidadBarras: Integer;
  iVeces: Smallint;
begin

  // Depuro las barras (\) de la normativa obtenida
  // Donde Hay 2 seguidas \\ dejo una \

  // Paso 2 veces (for) porque aveces hay 3 barras y debe dejar 1
  for iVeces := 1 to 2 do
  begin
    iCantidadBarras:=0;
    iItemCaracter:=1;
    while iItemCaracter<(Length(pTexto)-1) do //mientras que no se llegue a la posición final de la cadena
    begin
      if pTexto[iItemCaracter]='\' then
      begin
        iCantidadBarras:= iCantidadBarras + 1;
        if iCantidadBarras=2 then
        begin
          pTexto:= Copy(pTexto, 1, iItemCaracter-1) + Copy(pTexto, iItemCaracter+1, length(pTexto)-iItemCaracter+1);
          iCantidadBarras:=0;
        end;
      end
      else
      begin
        iCantidadBarras:=0;
      end;
      iItemCaracter:= iItemCaracter + 1;
    end; // while
  end; // for

  pTexto:= AnsiReplaceStr(pTexto,'\r\n',#13#10);

  Result:= pTexto;

end;

function TVSComunicacionParaReplicacion01._Descargar_Archivo_De_Servidor_FTP(
  pRutaLocal, pArchivo: String): Boolean;
begin
  Result := False;
  if IdFTP1.Connected then
  begin
    IdFTP1.TransferType := IdFTPCommon.TIdFTPTransferType.ftASCII;
    Try
      IdFTP1.Get(pArchivo, pRutaLocal+pArchivo, True);
      if FileExists(pRutaLocal+pArchivo) then
        Result := True;
    Except
      on EError: Exception do
        _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ 'No se pudo descargar el archivo ' + pArchivo + ' Error: ' + EError.Message);
    end;
  end
  else
    _Grabar_LogErrores_Comunicacion(-3, FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)+ 'No se pudo obtener archivo ' + pArchivo + ' falló conexion con el Servidor');
end;

procedure TVSComunicacionParaReplicacion01._Desconectar_Servidor_FTP;
begin
  if IdFTP1.Connected then
  begin
    IdFTP1.Disconnect;
  end;
end;

{function TVSComunicacionParaReplicacion1._Log_Procesado_En_Destino_Es_Mayor_Que_Local(
  pId_sucursal: Integer; pId_archivo_log_procesado: String; pUltimo_Pos: Integer): Boolean;
var
  iNro_Extension_Local, iNro_Extension_Destino: Integer;
begin
  Result:= False;
  with qrySelectLogProcesadoOrigenEnDestino do
  begin
    Parameters.ParamByName('pId_sucursal').Value := pId_sucursal;
    Open;
    if RecordCount > 0 then
    begin
      if not ((Trim(pId_archivo_log_procesado)=Trim(FieldByName('id_archivo_log_procesado_origen').AsString)) and (pUltimo_Pos=FieldByName('pos_log_procesado_origen').AsInteger)) then
      begin
        iNro_Extension_Local:= 0;
        iNro_Extension_Destino:= 0;
        try
          iNro_Extension_Local:= StrToInt(RightStr(pId_archivo_log_procesado, oParametro.Comu_largo_extension_file_binlog));
          iNro_Extension_Destino:= StrToInt(RightStr(Trim(FieldByName('id_archivo_log_procesado_origen').AsString), oParametro.Comu_largo_extension_file_binlog));
        except
          iNro_Extension_Local:= 0;
          iNro_Extension_Destino:= 0;
        end;

        if iNro_Extension_Destino > iNro_Extension_Local then
        begin
          Result:= True;
        end
        else
        begin
          if iNro_Extension_Destino = iNro_Extension_Local then
          begin
            if FieldByName('pos_log_procesado_origen').AsInteger > pUltimo_Pos then
            begin
              Result:= True;
            end;
          end;
        end;
      end;
    end;
    close;
  end;//end with qrySelectLogProcesadoOrigenEnDestino do
end;
}


function TVSComunicacionParaReplicacion01._Existe_log_Posterior_Destino(
  pId_archivo_log_procesado: String): Boolean;
var
  iNro_Extension: Integer;
begin
  iNro_Extension:= StrToInt(RightStr(pId_archivo_log_procesado, oParametro.Comu_largo_extension_file_binlog)) + 1;

  if Self.Tag<>100 then
    pId_archivo_log_procesado:= oParametro.Comu_nombre_file_binlog + '.'+ RightStr('00000000000000000'+IntToStr(iNro_Extension), oParametro.Comu_largo_extension_file_binlog)
  else
    // En AWS: mysql-bin-changelog - No se puede configurar
    pId_archivo_log_procesado:= 'mysql-bin-changelog' + '.'+ RightStr('00000000000000000'+IntToStr(iNro_Extension), oParametro.Comu_largo_extension_file_binlog);


  with qrySelect_Binlog_events_Destino_EPosterior do
  begin
    if qrySelect_Binlog_events_Destino_EPosterior.Active then
      qrySelect_Binlog_events_Destino_EPosterior.Close;
    //sql.Clear;
    //sql.Add('SHOW BINLOG EVENTS IN '+ QuotedStr(pId_archivo_log_procesado)+' LIMIT 10');
    //sql.Strings[0]:='SELECT * FROM bin_log_events_tabla WHERE bin_log_events_tabla.Log_name= '+ QuotedStr(pId_archivo_log_procesado) +' LIMIT 10';
    qrySelect_Binlog_events_Destino_EPosterior.Parameters.ParamByName('pLog_name').Value  := pId_archivo_log_procesado;
    try
      //Conexion_Remota.Connected:= False;
      //Conexion_Remota.Connected:= True;
      Open;
      if RecordCount > 1 then
        Result:= True
      else
        Result:= False;
    except
      Result:= False;
    end;
    close;
  end;//end with qrySelect_Binlog_events_Destino_EPosterior do
end;


function TVSComunicacionParaReplicacion01.CorrijeCadenaParticular(
  pCadena: String): String;
const
  //Caracteres_Validos = ['0'..'9', 'a'..'z', 'A'..'Z', chr(33), chr(34), chr(35), chr(36), chr(37), chr(38), chr(39), chr(40), chr(41), chr(42), chr(43), chr(44), chr(45), chr(46), chr(47), chr(58), chr(59), chr(60), chr(61), chr(62), chr(64), chr(123), chr(124), chr(125) ];
  Caracteres_Validos1 = ['0'..'9', 'a'..'z', 'A'..'Z'];
  Caracteres_Validos2 = [chr(32), chr(34),chr(35),chr(39),chr(40),chr(41),chr(42),chr(43),chr(44),chr(45),chr(46),chr(47),chr(58),chr(59),chr(60),chr(61),chr(62),chr(64)];

  {
  chr(32)     // espacio en blanco
  chr(34) "
  chr(35) #
  chr(39) '
  chr(40) (
  chr(41) )
  chr(42) *
  chr(43) +
  chr(44) ,
  chr(45) -
  chr(46) .
  chr(47) /
  chr(58) :
  chr(59) ;
  chr(60) <
  chr(61) =
  chr(62) >
  chr(64) @
  }
var
   iLoop,
   iMaxLoop  : Integer;
   sCaracter: Char;
begin
  iMaxLoop:= Length(pCadena);
  iLoop := 1;
  Result := '';
  while (iLoop <= iMaxLoop) do
  begin
    sCaracter:= pCadena[iLoop];
    {case sCaracter of
      //BLD: Result := (Result + 'A');
      'À','Á','Â','Ã','Ä','Å': sCaracter:= ' ';
      'à','á','â','ã','ä','å': sCaracter:= ' ';
      'È','É','Ê','Ë':         sCaracter:= ' ';
      'è','é','ê','ë':         sCaracter := ' ';
      'Ì','Í','Î','Ï':         sCaracter := ' ';
      'ì','í','î','ï':         sCaracter := ' ';
      'Ò','Ó','Ô','Õ','Ö':     sCaracter := ' ';
      'ò','ó','ô','õ','ö':     sCaracter := ' ';
      'Ù','Ú','Û','Ü':         sCaracter := ' ';
      'ù','ú','û','ü':         sCaracter := ' ';
      'Ñ':                     sCaracter := ' ';
      'ñ':                     sCaracter := ' ';
      'Ý':                     sCaracter := ' ';
      'ý':                     sCaracter := ' ';
      '|','~','¢','£','¤','¥',
      'º','»','¼','½','¾','¿',
      'Æ','Ç','Ð','Ø','Þ',
      'ß','æ','ç','ð','ø',
      'þ','ÿ','©',
      }
      { Agregados }
     { '•','¶','§','^','`',
      '×','ƒ','ª','®','¬',
      '¯','µ','´','«','‘',
      '±','÷','¸','°','¨',
      '·','¹','³','²','€','š','?': sCaracter := ' ';
      else                     sCaracter := sCaracter;
    end;}
    {
    case ord(sCaracter) of
      //'À','Á','Â','Ã','Ä','Å': sCaracter := '_';
      160,131,132,133,134:     sCaracter := ' ';
      //'È','É','Ê','Ë':         sCaracter := '_';
      130,136,137,138:         sCaracter := ' ';
      //'Ì','Í','Î','Ï':         sCaracter := '_';
      161,139,140,141:         sCaracter := ' ';
      //'Ò','Ó','Ô','Õ','Ö':     sCaracter := '_';
      162,147,148,149:         sCaracter := ' ';
      //'Ù','Ú','Û','Ü':         sCaracter := '_';
      163,150,151,129:         sCaracter := ' ';
      165:                     sCaracter := ' ';
      164:                     sCaracter := ' ';
      //'Ý':                     sCaracter := '_';
      //'ý':                     sCaracter := '_';
      166,167,168,169,170,171,172,173,174,175,176,177,178,179,180: sCaracter := ' ';
      184,185,186,187,188,189,190,191,192,193,194,195,196,197    : sCaracter := ' ';
      200,201,202,203,204,205,206,207,208,209                    : sCaracter := ' ';
      213                                                        : sCaracter := ' ';
      217,218,219,220,221                                        : sCaracter := ' ';
      223                                                        : sCaracter := ' ';
      225                                                        : sCaracter := ' ';
      230,231,232                                                : sCaracter := ' ';
      238,239,240,241,242,243,244,245,246,247,248,249,250,251,252: sCaracter := ' ';
      253,254                                                    : sCaracter := ' ';
      63                                                         : sCaracter := ' ';
      else                     sCaracter := sCaracter;
    end;
    }

    if ((sCaracter in Caracteres_Validos1) or (sCaracter in Caracteres_Validos2)) then
    begin
      sCaracter := sCaracter;
    end
    else
    begin
      sCaracter:= '#';
    end;

    Result:= Result + sCaracter;
    iLoop := (iLoop + 1);
  end;
  //Result:= AnsiReplaceStr(Result,'#','');
  Result:= AnsiReplaceStr(Result,'#',' ');
end;

end.

//****Servicio y unidades relacionadas desarrollados por Velia - Pablo ****//
