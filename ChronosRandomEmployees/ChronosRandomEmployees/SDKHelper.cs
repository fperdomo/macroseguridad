using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChronosRandomEmployees
{
    public class SDKHelper
    {
        public zkemkeeper.CZKEM axCZKEM1 = new zkemkeeper.CZKEM();

        private static bool bIsConnected = false;//the boolean value identifies whether the device is connected
        private static int iMachineNumber = 1;
        private static int idwErrorCode = 0;
        private static int iDeviceTpye = 1;
        bool bAddControl = true;        //Get all user's ID


        public bool GetConnectState()
        {
            return bIsConnected;
        }

        public void SetConnectState(bool state)
        {
            bIsConnected = state;
            //connected = state;
        }
        
        public int sta_ConnectTCP(List<string> lblOutputInfo, string ip, string port, string commKey)
        {
            if (ip == "" || port == "" || commKey == "")
            {
                lblOutputInfo.Add("*Name, IP, Port or Commkey cannot be null !");
                return -1;// ip or port is null
            }

            if (Convert.ToInt32(port) <= 0 || Convert.ToInt32(port) > 65535)
            {
                lblOutputInfo.Add("*Port illegal!");
                return -1;
            }

            if (Convert.ToInt32(commKey) < 0 || Convert.ToInt32(commKey) > 999999)
            {
                lblOutputInfo.Add("*CommKey illegal!");
                return -1;
            }

            int idwErrorCode = 0;

            axCZKEM1.SetCommPassword(Convert.ToInt32(commKey));

            if (bIsConnected == true)
            {
                axCZKEM1.Disconnect();              
                SetConnectState(false);
                lblOutputInfo.Add("Disconnect with device !");
                //connected = false;
                return -2; //disconnect
            }

            if (axCZKEM1.Connect_Net(ip, Convert.ToInt32(port)) == true)
            {
                SetConnectState(true);                
                lblOutputInfo.Add("Connect with device !");
                

                return 1;
            }
            else
            {
                axCZKEM1.GetLastError(ref idwErrorCode);
                lblOutputInfo.Add("*Unable to connect the device,ErrorCode=" + idwErrorCode.ToString());
                return idwErrorCode;
            }
        }

        public void sta_DisConnect()
        {
            if (GetConnectState() == true)
            {
                axCZKEM1.Disconnect();                
            }
        }

        public int sta_SetSMS(List<string> lblOutputInfo, string txtSMSID, string cbTag, string txtValidMin,
           string txtContent)
        {
            if (GetConnectState() == false)
            {
                lblOutputInfo.Add("*Please connect first!");
                return -1024;
            }

            if (txtSMSID.Trim() == "" || cbTag.Trim() == "" || txtValidMin.Trim() == "" || txtContent.Trim() == "")
            {
                lblOutputInfo.Add("*Please input data first!");
                return -1023;
            }

            if (Convert.ToInt32(txtSMSID.Trim()) <= 0)
            {
                lblOutputInfo.Add("*SMS ID error!");
                return -1023;
            }

            if (Convert.ToInt32(txtValidMin.Trim()) < 0 || Convert.ToInt32(txtValidMin.Trim()) > 65535)
            {
                lblOutputInfo.Add("*Expired time error!");
                return -1023;
            }

            int idwErrorCode = 0;
            int iSMSID = Convert.ToInt32(txtSMSID.Trim());
            int iTag = 0;
            int iValidMins = Convert.ToInt32(txtValidMin.Trim());
            string sStartTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm").Trim();
            string sContent = txtContent.Trim();
            string sTag = cbTag.Trim();

            for (iTag = 253; iTag <= 255; iTag++)
            {
                if (sTag.IndexOf(iTag.ToString()) > -1)
                {
                    break;
                }
            }

            axCZKEM1.EnableDevice(iMachineNumber, false);
            if (axCZKEM1.SetSMS(iMachineNumber, iSMSID, iTag, iValidMins, sStartTime, sContent))
            {
                axCZKEM1.RefreshData(iMachineNumber);//After you have set the short message,you should refresh the data of the device
                lblOutputInfo.Add("Successfully set SMS! SMSType=" + iTag.ToString());
            }
            else
            {
                axCZKEM1.GetLastError(ref idwErrorCode);
                lblOutputInfo.Add("*Operation failed,ErrorCode=" + idwErrorCode.ToString());
            }
            axCZKEM1.EnableDevice(iMachineNumber, true);

            return idwErrorCode != 0 ? idwErrorCode : 1;
        }

        public int sta_SetUserSMS(List<string> lblOutputInfo, int txtSMSID, long cbUserID)
        {
            if (GetConnectState() == false)
            {
                lblOutputInfo.Add("*Please connect first!");
                return -1024;
            }


            int idwErrorCode = 0;
            int iSMSID = txtSMSID;
            int iTag = 0;
            int iValidMins = 0;
            string sStartTime = "";
            string sContent = "";
            string sEnrollNumber = cbUserID.ToString();

            axCZKEM1.EnableDevice(iMachineNumber, false);

            if (axCZKEM1.GetSMS(iMachineNumber, iSMSID, ref iTag, ref iValidMins, ref sStartTime, ref sContent) == false)
            {
                lblOutputInfo.Add("*The SMSID doesn't exist!!");
                axCZKEM1.EnableDevice(iMachineNumber, true);
                return -1022;
            }

            if (iTag != 254)
            {
                lblOutputInfo.Add("*The SMS does not Personal SMS,please set it as Personal SMS first!!");
                axCZKEM1.EnableDevice(iMachineNumber, true);
                return -1022;
            }

            if (axCZKEM1.SSR_SetUserSMS(iMachineNumber, sEnrollNumber, iSMSID))
            {
                axCZKEM1.RefreshData(iMachineNumber);//After you have set user short message,you should refresh the data of the device
                lblOutputInfo.Add("Successfully set user SMS! ");
            }
            else
            {
                axCZKEM1.GetLastError(ref idwErrorCode);
                lblOutputInfo.Add("*Operation failed,ErrorCode=" + idwErrorCode.ToString());
            }
            axCZKEM1.EnableDevice(iMachineNumber, true);

            return idwErrorCode != 0 ? idwErrorCode : 1;
        }
    }
}
