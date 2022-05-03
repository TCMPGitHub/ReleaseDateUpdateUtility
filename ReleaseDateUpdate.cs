using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;

namespace ReleaseDateUpdateUtility
{
    public class ReleaseDateUpdate 
    {
        public string DBBassConn { get; set; }
        public string DBPatsConn { get; set; }

        public string FilePath { get; set; }
        public string ErrorMessage { get; set; }

        //public event EventHandler Disposed;
        public string  Import()
        {
            LogWriter.LogMessageToFile("Get all CDCRs from import file.");
            //string excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\\DHCSDates\\Q42017Results.xlsx;Extended Properties=\"Excel 12.0;HDR=YES;\"";
            //string CSVFileConnectionString = String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"text;HDR=Yes;FMT=Delimited\";", Path.GetDirectoryName(FilePath));
            string CSVFileConnectionString = String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"text;HDR=Yes;FMT=Delimited\";", Path.GetDirectoryName(FilePath));
            try
            {
                DataTable dt = new DataTable();
                using (OleDbConnection con = new OleDbConnection(CSVFileConnectionString))
                {
                    con.Open();
                    var csvQuery = string.Format("select [CDCNUMBER],[SCHEDULEDRELEASEDATE] from [{0}]", Path.GetFileName(FilePath));
                    //using (OleDbDataAdapter da = new OleDbDataAdapter(csvQuery, con))
                        
                //using (OleDbConnection con = new OleDbConnection(CSVFileConnectionString))
                //{
                //    con.Open();
                    //var csvQuery = string.Format("select [CDCNUMBER],[SCHEDULEDRELEASEDATE] from[{0}]", Path.GetFileName(FilePath));
                    using (OleDbDataAdapter da = new OleDbDataAdapter(csvQuery, con))
                    {
                        try
                        {
                            da.Fill(dt);
                        }
                        catch (Exception ex)
                        {
                            LogWriter.LogMessageToFile(ex.Message);
                        }
                    }

                    dt.TableName = "ReleaseDateChanges";
                }

                LogWriter.LogMessageToFile("Read data from file Complated");

                if (UpdateReleaseDates(dt))
                {
                    LogWriter.LogMessageToFile("Total " + dt.Rows.Count + " CDCRs Updated.");
                    return "Total " + dt.Rows.Count + " CDCRs Updated.";
                }
                else
                {
                    LogWriter.LogMessageToFile("Failed to update CDCRs.");
                    return "Failed to update CDCRs.";
                }
            }
            catch (Exception ex)
            {
                LogWriter.LogMessageToFile(ex.Message);
                return ex.Message;
            }
        }

        private bool UpdateReleaseDates(DataTable dt)
        {
            try
            {
                //Update DHCSDates only for Medical
                using (SqlConnection cnz = new SqlConnection(DBBassConn))
                {
                    try
                    {
                        cnz.Open();
                        SqlCommand cmdzs = new SqlCommand();
                        SqlParameter parameter = new SqlParameter();
                        cmdzs.Connection = cnz;
                        cmdzs.Parameters.AddWithValue("@ReleaseDateChanges", dt);
                        cmdzs.CommandText = "spUpdateReleaseDates";
                        cmdzs.CommandType = CommandType.StoredProcedure;
                        cmdzs.CommandTimeout = 300;
                        cmdzs.ExecuteNonQuery();
                        return true;
                    }
                    catch (SqlException err)
                    {
                        LogWriter.LogMessageToFile(err.Message);
                    }
                    finally
                    {
                        cnz.Close();
                    }
                    return false;
                }
            }
            catch (Exception ex)
            {
                LogWriter.LogMessageToFile(ex.Message);
                return false;
            }

        }

        //public int ImportDate()
        //{

        //    DataTable dt = new DataTable("CSVTable");
        //    LogWriter.LogMessageToFile("Read dates from file: ");

        //    string CSVFileConnectionString = String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"text;HDR=Yes;FMT=Delimited\";", Path.GetDirectoryName(FilePath));
        //    using (OleDbConnection con = new OleDbConnection(CSVFileConnectionString))
        //    {
        //        con.Open();
        //        var csvQuery = string.Format("select * from [{0}]", Path.GetFileName(FilePath));
        //        using (OleDbDataAdapter da = new OleDbDataAdapter(csvQuery, con))
        //        {
        //            da.Fill(dt);

        //        }
        //        if (dt.Rows.Count > 0)
        //        {
        //            using (SqlBulkCopy bulkcopy = new SqlBulkCopy(DBBassConn))
        //            {
        //                bulkcopy.ColumnMappings.Add("CDCNUMBER", "CDCNUMBER");
        //                bulkcopy.ColumnMappings.Add("SCHEDULEDRELEASEDATE", "SCHEDULEDRELEASEDATE");

        //                bulkcopy.DestinationTableName = "ReleaseDateChanges";
        //                bulkcopy.BatchSize = 0;
        //                bulkcopy.WriteToServer(dt);
        //                bulkcopy.Close();
        //            }
        //        }
        //    }
        //    if (dt.Rows.Count > 0)
        //    {
        //        var query = @"UPDATE e SET e.[ReleaseDate]= r.[SCHEDULEDRELEASEDATE] FROM [BassWeb].[dbo].[Episode] AS e 
        //                               INNER JOIN [BassWebTest].[dbo].[ReleaseDateChanges] AS r ON e.[CDCRNum] = r.[CDCNUMBER]
        //                      UPDATE e SET e.[ReleaseDate]= r.[SCHEDULEDRELEASEDATE] FROM [PatsWebV2].[dbo].[Episode] AS e 
        //                               INNER JOIN [BassWebTest].[dbo].[ReleaseDateChanges] AS r ON e.[CDCRNum] = r.[CDCNUMBER]
        //                      TRUNCATE TABLE [BassWebTest].[dbo].[ReleaseDateChanges]";
        //        try
        //        {
        //            using (SqlConnection conn = new SqlConnection(DBBassConn))
        //            {
        //                conn.Open();
        //                SqlCommand command = new SqlCommand(query, conn);
        //                command.CommandTimeout = 3000;
        //                command.ExecuteNonQuery();
        //            }
        //            return dt.Rows.Count;
        //        }
        //        catch (Exception ex) { } // Handle exception properly           
        //    }
        //    return 0;
        //}

        //protected virtual void Dispose(bool disposing)
        //{
        //    if (disposing)
        //    {
        //        // Free other state (managed objects).
        //    }
        //}

        //public void Dispose()
        //{
        //    throw new NotImplementedException();
        //}

        ~ReleaseDateUpdate() { }
        //{
        //    Dispose(false);
        //}
    }
}
