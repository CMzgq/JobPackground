   /// <summary>   
   
        #region
        /// 根据excel 路径把第一个sheet中的内容放入到datatable
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        ///      
        
        
          public static DataTable ReadExcelToTable(string path)
         {
            try
            {             
                //  连接字符串
                string connstring = "Provider=Microsoft.ACE.OleDb.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=0'"; 
                using (OleDbConnection conn = new OleDbConnection(connstring)) 
                {
                    conn.Open();
                    DataTable sheetsName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { 
                      null,null,null,"Table"
                    });     //得到所有sheet的名字
                    string firstsheetName = sheetsName.Rows[0][2].ToString();//得到第一个sheet的名字
                     
                    string cmdText = string.Format("select * from [{0}]", firstsheetName);//查询字符串
                    OleDbDataAdapter ada = new OleDbDataAdapter(cmdText, conn);
                    DataSet ds = new DataSet();
                    ada.Fill(ds);
                    return  ds.Tables[0];
                        

                }
            }
            catch (Exception)
            {

                return null;
            }          
        }

         
         #endregion
        
      
