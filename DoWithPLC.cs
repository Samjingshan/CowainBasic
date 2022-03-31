using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using PLC_Adapt_;

namespace Sharetek
{
   public class DoWithPLC
    {
    
        Dictionary<String, String> _DicPositionPort = new Dictionary<string, string>();
       
        PLCManager plcManager = null;
        public DoWithPLC(PLCManager PM)
        {
            this.plcManager = PM as PLCManager;
            this.plcManager.PLCEvent += LoaderPlcManager_PLCEvent;

              
            
        }



        

        private void LoaderPlcManager_PLCEvent(PLCEvent Event)
        {
            try
            {
                if (Event.item != null)
                {
                    switch (Event.item.name)
                    {
                        case "Read_Code_Flag":
                            {
                               
                               

                                if (Event.newValue != "0")//Read_Code_Flag由0置1
                                {
                                   
                                }
                                break;
                            }
                    }
                }
            }
            catch (Exception ex)
            {
                //_Message = CoreHelper.GetExceptionInfo(ex).Message;
                //ShowMessage(_Message);
                //CoreHelper.Logger.Error(_Message);
            }
        }

        




        
      
        
      
        
    }

}
