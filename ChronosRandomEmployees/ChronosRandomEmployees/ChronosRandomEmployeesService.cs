using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace ChronosRandomEmployees
{
    public partial class ChronosRandomEmployeesService : ServiceBase
    {

        public SDKHelper SDK = new SDKHelper();
        string _dsnName;
        string _dsnUser;
        string _dsnPass;
        string _timeExecuteStr;
        string _countEmpStr;
        int _countEmp;
        string _message;
        List<Device> _devicesReintent;
        Timer timer = new Timer();
        public ChronosRandomEmployeesService()
        {
            InitializeComponent();
        }
        protected override void OnStart(string[] args)
        {
            GetConfiguration();

            timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
            timer.Interval = 300000; //number in milisecinds, execute per 5 minute
            timer.Enabled = true;
        }

        private void GetConfiguration()
        {
            try
            {
                _devicesReintent = new List<Device>();
                Configuration configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                AppSettingsSection appSettings = configuration.AppSettings;
                _dsnName = appSettings.Settings["ODBC"].Value;
                _dsnUser = appSettings.Settings["ODBCUser"].Value;
                _dsnPass = appSettings.Settings["ODBCPass"].Value;
                _timeExecuteStr = appSettings.Settings["HoraEjecucion"].Value;
                _countEmpStr = appSettings.Settings["CantidadEmpleados"].Value;
                _message = appSettings.Settings["Mensaje"].Value;

                if (!int.TryParse(_countEmpStr, out _countEmp))
                    WriteToFile("Error obteniendo CantidadEmpleados de configuración" + DateTime.Now);

                WriteToFile("Iniciando el servicio ... " + DateTime.Now);
                WriteToFile("----------------- Configuración -----------------");
                WriteToFile("ODBC: " + _dsnName);
                WriteToFile("ODBCUser: " + _dsnUser);
                WriteToFile("ODBCPass: " + _dsnPass);
                WriteToFile("HoraEjecucion: " + _timeExecuteStr);
                WriteToFile("CantidadEmpleados: " + _countEmp);
                WriteToFile("Mensaje: " + _message);
                WriteToFile("------------------------------------------------");
            }
            catch (Exception ex)
            {
                WriteToFile("ERROR: Error obteniendo configuración" + " - " + ex.Message);
            }

        }
        protected override void OnStop()
        {
            WriteToFile("ERROR: Se detuvo el servicio");
        }

        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {           
            try
            {
                DateTime now = DateTime.Now;
                DateTime? lastDateExecute = (DateTime?)null;
                List<long> employees = new List<long>();
                WriteToFile("Comienza ejecución de Random");
                OdbcConnection cnn = OpenConnect();
                DateTime dateExecute = GetDateExecute();                             
                try
                {
                    bool executeNow = false;
                    lastDateExecute = GetLastExecuteService(cnn);
                    if (lastDateExecute.HasValue)
                    {
                        WriteToFile("Ultima Ejecución: " + lastDateExecute.Value.ToString());
                        if (lastDateExecute.Value.AddHours(24) < dateExecute)
                            executeNow = true;
                    }                        
                    else
                    {
                        WriteToFile("Primera Ejecución del Servicio");
                        if (now.AddMinutes(-6) < dateExecute && dateExecute < now.AddMinutes(6))
                            executeNow = true;
                    }
                        

                    if (executeNow)
                    {
                        if (lastDateExecute.HasValue)
                        {
                            WriteToFile("Borrar mensaje a los usuarios anteriores...");
                            ClearUserMessage(cnn, lastDateExecute.Value);
                        }
                        

                        WriteToFile("Obtener los empleados de la BD");
                        employees = GetActiveEmployees(cnn);

                        if (employees.Any())
                        {
                            Random rnd = new Random();
                            List<long> selectedEmployees = employees.OrderBy(x => rnd.Next()).Take(_countEmp).ToList();
                            foreach (var employeeId in selectedEmployees)
                            {
                                InsertSelectedEmployee(cnn, employeeId, dateExecute);
                            }
                            List<Device> totalDevices = GetDevices(cnn);                            
                            SendDataToDevices(cnn, selectedEmployees, totalDevices);
                            InsertLastExecuteService(cnn, dateExecute, !lastDateExecute.HasValue);
                        }
                        else
                        {
                            WriteToFile("WARNING: No se obtuvieron empleados de la base.");
                        }
                    }
                    else
                    {
                        WriteToFile("INFO: No es tiempo de ejecutarse: " + dateExecute);
                        if (lastDateExecute.HasValue)
                        {
                            WriteToFile("INFO: REINTENTO envio de: " + lastDateExecute);                         
                            var selectedEmployees = GetSelectedEmployee(cnn, lastDateExecute.Value);
                            SendDataToDevices(cnn, selectedEmployees, _devicesReintent);
                        }                      
                    }
                }
                catch (Exception ex)
                {
                    WriteToFile("ERROR: Error al obtener datos de la base. " + " - " + ex.Message);
                }
                finally
                {
                    cnn.Close();
                }
            }
            catch (Exception ex)
            {
                WriteToFile(ex.Message);
            }
        }
        private DateTime GetDateExecute()
        {
            DateTime result = DateTime.Now;
            try
            {
                var splitTimeExecute = _timeExecuteStr.Split(':');
                int hour = result.Hour;
                int minute = result.Minute;
                int.TryParse(splitTimeExecute[0], out hour);
                int.TryParse(splitTimeExecute[1], out minute);
                var now = DateTime.Now;
                result = new DateTime(now.Year, now.Month, now.Day, hour, minute, 0);
            }
            catch (Exception ex)
            {
                WriteToFile("ERROR: Error al crear la fecha de ejecución, verificar App.config " + " - " + ex.Message);
            }
            return result;
        }

        private bool ShowErrors(List<string> lstErrors)
        {
            foreach (var error in lstErrors)
            {
                WriteToFile("ERROR: Error con el reloj: " + error);               
            }
            return false;
        }

        private bool SendDataToDevices(OdbcConnection cnn, IEnumerable<long> selectedEmployees, List<Device> devices)
        {
            List<Device> deviceError = new List<Device>();
            bool resultOk = true;            
            WriteToFile("Comienza envio a relojes");
            try
            {
                foreach (var device in devices)
                {
                    List<string> lstErrorsConn = new List<string>();
                    List<string> lstErrorsSMS = new List<string>();
                    List<string> lstErrorsSMSUsers = new List<string>();

                    WriteToFile(string.Format("Conectando al reloj Id:{0}, Ip:{1}, Puerto:{2}", device.Id, device.Ip, device.Puerto));
                    int connectionResult = SDK.sta_ConnectTCP(lstErrorsConn, device.Ip, "4370", "0");
                    if (connectionResult <= 0)
                    {
                        resultOk = ShowErrors(lstErrorsConn);
                        if (!deviceError.Contains(device))
                            deviceError.Add(device);
                    }                        
                    else
                    {
                        WriteToFile("Conexión con reloj - OK");
                        WriteToFile("Enviando Mensaje");
                        int sendSMSResult = SDK.sta_SetSMS(lstErrorsSMS, 1.ToString(), 254.ToString(), 1440.ToString(), _message); ;
                        if (sendSMSResult <= 0)
                        {
                            resultOk = ShowErrors(lstErrorsSMS);
                            if (!deviceError.Contains(device))
                                deviceError.Add(device);
                        }                            
                        else
                        {
                            WriteToFile("Mensaje Enviado - OK");
                            WriteToFile("Empleados Seleccionados");
                            foreach (var employeeId in selectedEmployees)
                            {
                                WriteToFile("Mensaje enviado a empleado Nro. " + employeeId + " Dispositivo Id:" + device.Ip);

                                int sendSMSUserResult = SDK.sta_SetUserSMS(lstErrorsSMSUsers, 1, employeeId);
                                if (sendSMSUserResult <= 0)
                                {
                                    resultOk = ShowErrors(lstErrorsSMSUsers);
                                    if (!deviceError.Contains(device))
                                        deviceError.Add(device);
                                }  
                            }
                        }
                        WriteToFile(string.Format("Desconectado reloj Id:{0}, Ip:{1}, Puerto:{2}", device.Id, device.Ip, device.Puerto));
                        SDK.sta_DisConnect();
                    }                    
                }
                _devicesReintent = deviceError;
            }
            catch (Exception ex)
            {
                WriteToFile("ERROR: Error conectando y enviando datos a los relojes: " + ex.Message);
            }
           
            return resultOk;
        }

        public void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(DateTime.Now + " - " + Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(DateTime.Now + " - " + Message);
                }
            }
        }

        private void ClearUserMessage(OdbcConnection cnn, DateTime dateExecute)
        {
            var selectedEmployees = GetSelectedEmployee(cnn, dateExecute);
            List<Device> totalDevices = GetDevices(cnn);
            ClearDataUserToDevices(cnn, selectedEmployees, totalDevices, dateExecute);
        }

        private bool ClearDataUserToDevices(OdbcConnection cnn, IEnumerable<long> selectedEmployees, List<Device> devices, DateTime dateExecute)
        {
            List<Device> deviceError = new List<Device>();
            bool resultOk = true;
            WriteToFile("L: Comienza limpieza de mensaje de usuarios en relojes");
            try
            {
                foreach (var device in devices)
                {
                    List<string> lstErrorsConn = new List<string>();
                    List<string> lstErrorsSMS = new List<string>();
                    List<string> lstErrorsSMSUsers = new List<string>();

                    WriteToFile(string.Format("Conectando al reloj Id:{0}, Ip:{1}, Puerto:{2}", device.Id, device.Ip, device.Puerto));
                    int connectionResult = SDK.sta_ConnectTCP(lstErrorsConn, device.Ip, "4370", "0");
                    if (connectionResult <= 0)
                    {
                        ShowErrors(lstErrorsConn);
                    }else
                    {
                        foreach (var employeeId in selectedEmployees)
                        {                            
                            int sendSMSUserResult = SDK.sta_DelUserSMS(lstErrorsSMSUsers, 1, employeeId);
                            if (sendSMSUserResult <= 0)
                            {
                                resultOk = ShowErrors(lstErrorsSMSUsers);                                
                            }
                            else
                            {
                                WriteToFile("Mensaje BORRADO a empleado Nro. " + employeeId + " Dispositivo Id:" + device.Ip);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteToFile("ERROR: Error conectando y vaciando mensaje de usuario en relojes: " + ex.Message);
            }

            return resultOk;
        }

        #region ConnectToDataBase

        private OdbcConnection OpenConnect()
        {
            OdbcConnection DbConnection = null;
            try
            {               
                 var stringConeccion = "dsn=" +
                _dsnName + ";uid=" +
                _dsnUser + ";pwd=" +
                _dsnPass + ";";
                DbConnection = new OdbcConnection(stringConeccion);
                DbConnection.Open();
                WriteToFile("Connectado a ODBC: " + _dsnName);
            }
            catch (OdbcException ex)
            {
                WriteToFile("ERROR: Error al conectar con ODBC. " + " - " + ex.Message);
            }
            return DbConnection;
        }

        private List<long> GetActiveEmployees(OdbcConnection cnn)
        {
            List<long> employees = new List<long>();
            OdbcCommand DbCommand;
            OdbcDataReader DbReader;
            try
            {
                DbCommand = cnn.CreateCommand();
                DbCommand.CommandText = "SELECT EmpleadoID FROM EMPLEADO WHERE EmpleadoInactivo = 0";
                DbReader = DbCommand.ExecuteReader();
                int fCount = DbReader.FieldCount;
                while (DbReader.Read())
                {
                    for (int i = 0; i < fCount; i++)
                    {
                        employees.Add(DbReader.GetInt32(i));
                    }
                }
                DbReader.Close();
                DbCommand.Dispose();
            }
            catch (OdbcException ex)
            {
                WriteToFile("ERROR: Error al obtener empleados de la base. " + " - " + ex.Message);
            }
            return employees;
        }

        private DateTime? GetLastExecuteService(OdbcConnection cnn)
        {
            DateTime? result = (DateTime?)null;
            OdbcCommand DbCommand;
            OdbcDataReader DbReader;
            try
            {
                DbCommand = cnn.CreateCommand();
                DbCommand.CommandText = "SELECT UltimoMarca FROM ULTIMO WHERE UltimoEmpleado = ?";
                DbCommand.Parameters.Add("@param1", OdbcType.Int).Value = "999999999";
                DbReader = DbCommand.ExecuteReader();
                int fCount = DbReader.FieldCount;
                while (DbReader.Read())
                {
                    for (int i = 0; i < fCount; i++)
                    {
                        result = (DbReader.GetDateTime(0));
                    }
                }
                DbReader.Close();
                DbCommand.Dispose();
            }
            catch (OdbcException ex)
            {
                WriteToFile("ERROR: Error al obtener la ultima ejecución de la base. " + " - " + ex.Message);
            }

            return result;
        }

        private void InsertLastExecuteService(OdbcConnection cnn, DateTime dateExecute, bool isFirstExecute)
        {
            OdbcCommand DbCommand;
            try
            {
                DbCommand = cnn.CreateCommand();
                if (isFirstExecute)
                {
                    DbCommand.CommandText = "INSERT INTO ULTIMO(UltimoEmpleado, UltimoMarca) VALUES (?, ?)";
                    DbCommand.Parameters.Add("@param1", OdbcType.Int).Value = "999999999";
                    DbCommand.Parameters.Add("@param2", OdbcType.DateTime).Value = dateExecute;
                }
                else
                {
                    DbCommand.CommandText = "UPDATE ULTIMO SET UltimoMarca = ? WHERE UltimoEmpleado = ?";
                    DbCommand.Parameters.Add("@param1", OdbcType.DateTime).Value = dateExecute;
                    DbCommand.Parameters.Add("@param2", OdbcType.Int).Value = "999999999";
                }              
               
                DbCommand.ExecuteNonQuery();

                DbCommand.Dispose();
            }
            catch (OdbcException ex)
            {
                WriteToFile("ERROR: Error al insertar la ultima ejecución de la base. " + " - " + ex.Message);
            }

        }
        
        private void InsertSelectedEmployee(OdbcConnection cnn, long employeeId, DateTime dateExecute)
        {
            OdbcCommand DbCommand;
            try
            {
                DbCommand = cnn.CreateCommand();
                DbCommand.CommandText = "INSERT INTO LIQHORAS1(EmpleadoID, LiqHoraFecha,LiqSeccionTrabaja) VALUES (?, ? ,?)";
                DbCommand.Parameters.Add("@param1", OdbcType.Int).Value = employeeId;
                DbCommand.Parameters.Add("@param2", OdbcType.DateTime).Value = dateExecute;
                DbCommand.Parameters.Add("@param3", OdbcType.Int).Value = "999999999";
                DbCommand.ExecuteNonQuery();

                DbCommand.Dispose();
            }
            catch (OdbcException ex)
            {
                WriteToFile("ERROR: Error al insertar los empleados seleccionados de la base. " + " - " + ex.Message);
            }
        }

        private IEnumerable<long> GetSelectedEmployee(OdbcConnection cnn, DateTime dateExecute)
        {
            List<long> result = new List<long>();
            OdbcCommand DbCommand;
            OdbcDataReader DbReader;
            try
            {
                DbCommand = cnn.CreateCommand();
                DbCommand.CommandText = "SELECT EmpleadoID FROM LIQHORAS1  WHERE LiqHoraFecha = ? AND LiqSeccionTrabaja = ?";
                DbCommand.Parameters.Add("@param1", OdbcType.DateTime).Value = dateExecute;
                DbCommand.Parameters.Add("@param2", OdbcType.Int).Value = "999999999";
                DbReader = DbCommand.ExecuteReader();
                int fCount = DbReader.FieldCount;
                while (DbReader.Read())
                {
                    for (int i = 0; i < fCount; i++)
                    {
                        result.Add(DbReader.GetInt64(0));
                    }
                }
                DbReader.Close();
                DbCommand.Dispose();
            }
            catch (OdbcException ex)
            {
                WriteToFile("ERROR: Error al obtener la ultima ejecución de la base. " + " - " + ex.Message);
            }

            return result;
        }

        private List<Device> GetDevices(OdbcConnection cnn)
        {
            List<Device> devices = new List<Device>();
            try
            {
                OdbcCommand DbCommand = cnn.CreateCommand();
                DbCommand.CommandText = "SELECT IdDispositivo,DirIP,Puerto FROM DISPOSITIVOS";
                OdbcDataReader DbReader = DbCommand.ExecuteReader();
                int fCount = DbReader.FieldCount;
                while (DbReader.Read())
                {
                    Device newDevice = new Device();
                    newDevice.Id = DbReader.GetInt32(0);
                    newDevice.Ip = DbReader.GetString(1).Trim();
                    newDevice.Puerto = DbReader.GetString(2).Trim();
                    devices.Add(newDevice);
                }
                DbReader.Close();
                DbCommand.Dispose();
            }
            catch (OdbcException ex)
            {
                WriteToFile("ERROR: Error al obtener dispositivos de la base. " + " - " + ex.Message);
            }
            return devices;
        }
        #endregion
    }
}

