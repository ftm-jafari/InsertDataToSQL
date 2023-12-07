
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using Google.Protobuf.WellKnownTypes;
using MySql.Data.MySqlClient;
using MySqlX.XDevAPI.Relational;
using OfficeOpenXml;
using static InsertDataToSQL.Program;

namespace InsertDataToSQL
{

    class Program
    {
        //****product****
        public class product
        {
            public string code_kala { get; set; }
            public string name_kala { get; set; }
            public string vahed { get; set; }
            public string kind { get; set; }
            public string disn { get; set; }
            public string color { get; set; }
            public string t_san { get; set; }
        }

        //****out-product****
        public class outProduct
        {
            public string code_kala { get; set; }
            public string no_c { get; set; }
            public string wnet { get; set; }
            public string nwnet { get; set; }
            public string n_dok { get; set; }
            public string k_dok { get; set; }
            public string hambaft { get; set; }
        }

        static HashSet<string> colorValue = new HashSet<string>();
        static HashSet<string> disnValue = new HashSet<string>();
        static HashSet<string> kindValue = new HashSet<string>();
        static HashSet<string> t_sanValue = new HashSet<string>();
        static HashSet<string> hambaftValue = new HashSet<string>();
        static HashSet<string> k_dokValue = new HashSet<string>();
        static HashSet<string> n_dokValue = new HashSet<string>();
        static HashSet<string> units = new HashSet<string>();


        static void Main()
        {
            string excelFilePath = "D:\\PishroTech\\InsertDataToSQL\\InsertDataToSQL\\data\\kala.xlsx";
            string excelFilePathSecound = "D:\\PishroTech\\InsertDataToSQL\\InsertDataToSQL\\data\\i_in.xlsx";
            string connectionString = "Server=localhost;Database=sadra;User ID=root;Password=;";

            DataTable excelDataKala = ReadExcelData(excelFilePath);
            DataTable excelDataOutProduct = ReadExcelData(excelFilePathSecound);


          


            //product
            var products = new List<product>();

            //out-product
            var outProducts = new List<outProduct>();


            if (excelDataKala != null)
            {
                foreach (DataRow item in excelDataKala.Rows)
                {
                    products.Add(new product()
                    {
                        code_kala = item["code_kala"].ToString(),
                        name_kala = item["name_kala"].ToString(),
                        vahed = item["vahed"].ToString(),
                        color = item["color"].ToString(),
                        disn = item["disn"].ToString(),
                        kind = item["kind"].ToString(),
                        t_san = item["t_san"].ToString(),
                    });
                }
            }


            if (excelDataOutProduct != null)
            {
                foreach (DataRow item in excelDataOutProduct.Rows)
                {
                    outProducts.Add(new outProduct()
                    {
                        code_kala = item["code_kala"].ToString(),
                        no_c = item["no_c"].ToString(),
                        wnet = item["wnet"].ToString(),
                        hambaft = item["hambaft"].ToString(),
                        k_dok = item["k_dok"].ToString(),
                        n_dok = item["n_dok"].ToString(),
                        nwnet = item["nwnet"].ToString()
                    });
                }
            }


            colorValue = new HashSet<string>(products.Select(x => x.color).ToList());
            disnValue = new HashSet<string>(products.Select(x => x.disn).ToList());
            kindValue = new HashSet<string>(products.Select(x => x.kind).ToList());
            t_sanValue = new HashSet<string>(products.Select(x => x.t_san).ToList());
            hambaftValue = new HashSet<string>(outProducts.Select(x => x.hambaft).ToList());
            k_dokValue = new HashSet<string>(outProducts.Select(x => x.k_dok).ToList());
            n_dokValue = new HashSet<string>(outProducts.Select(x => x.n_dok).ToList());

            units = new HashSet<string>(products.Select(x => x.vahed).ToList());


            // اضافه شدن مرحله تولید
            var phaseId = InsertPhasesIntoMySql(connectionString);

            // اضافه شدن گروه مرحله تولید
            var profileId = InsertProfileIntoMySql(connectionString);
            InsertBuildPhaseProfileIntoMySql(new List<int>() { phaseId, profileId }, connectionString);

            // اضافه شدن پارامتر / مقادیر مرحله تولید
            InsertBuildPhaseParametersIntoMySql(phaseId, connectionString);

            // اضافه شدن واحد
            InsertUnitIntoMySql(units, connectionString);
            InsertProductIntoMySql(products, profileId, connectionString);
        }


        static DataTable ReadExcelData(string filePath)
        {
            using (var package = new ExcelPackage(new System.IO.FileInfo(filePath)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                var workSheet = package.Workbook.Worksheets[0];
                var table = new DataTable();

                // Assume the first row contains the column names
                foreach (var firstRowCell in workSheet.Cells[1, 1, 1, workSheet.Dimension.End.Column])
                {
                    table.Columns.Add(firstRowCell.Text.Trim());
                }

                // Start adding data from the second row
                for (var rowNumber = 2; rowNumber <= workSheet.Dimension.End.Row; rowNumber++)
                {
                    var row = workSheet.Cells[rowNumber, 1, rowNumber, workSheet.Dimension.End.Column];
                    var newRow = table.Rows.Add();
                    foreach (var cell in row)
                    {
                        newRow[cell.Start.Column - 1] = cell.Text.Trim();
                    }
                }

                return table;
            }
        }

        //bulid-phases مرحله تولید 
        static int InsertPhasesIntoMySql(string connectionString)
        {

            var phaseId = 0;

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();

                var query = $"INSERT INTO build_phases (`title`, `type`, `order`, `in_warehouse`) VALUES(@Title, @Type, @Order, @In_Warehouse)";

                using MySqlCommand command = new MySqlCommand(query, connection);
                command.Parameters.AddWithValue($"@Title", "تولید");
                command.Parameters.AddWithValue($"@Type", "packing");
                command.Parameters.AddWithValue($"@Order", 1);
                command.Parameters.AddWithValue($"@In_Warehouse", 0);
                command.ExecuteNonQuery();

                var getPhasesQuery = "SELECT * FROM build_phases WHERE title = @Title";
                using (MySqlCommand getPhsesCommand = new MySqlCommand(getPhasesQuery, connection))
                {
                    getPhsesCommand.Parameters.AddWithValue("@Title", "تولید");

                    using (MySqlDataReader reader = getPhsesCommand.ExecuteReader())
                    {
                        reader.Read();
                        phaseId = reader.GetInt32("id");
                        reader.Close();

                    }
                }

                connection.Close();

            }
            Console.WriteLine("Adding Phases was successful!");
            return phaseId;
        }

        //phase_profiles پروفایل تولید 
        static int InsertProfileIntoMySql(string connectionString)
        {

            var profileId = 0;

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();

                var query = $"INSERT INTO phase_profiles (`title`, `phases`) VALUES(@Title, @Phases)";

                using MySqlCommand command = new MySqlCommand(query, connection);
                command.Parameters.AddWithValue($"@Title", "تولید");
                command.Parameters.AddWithValue($"@Phases", "[1]");
                command.ExecuteNonQuery();

                var getProfileQuery = "SELECT * FROM phase_profiles WHERE title = @Title";
                using (MySqlCommand getProfileCommand = new MySqlCommand(getProfileQuery, connection))
                {
                    getProfileCommand.Parameters.AddWithValue("@Title", "تولید");

                    using (MySqlDataReader reader = getProfileCommand.ExecuteReader())
                    {
                        reader.Read();
                        profileId = reader.GetInt32("id");
                        reader.Close();

                    }
                }

                connection.Close();

            }
            Console.WriteLine("Adding Profile was successful!");
            return profileId;
        }

        //build_phase_phase_profile ارتباط مرحله و پروفایل
        static void InsertBuildPhaseProfileIntoMySql(List<int> ids, string connectionString)
        {

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();

                var query = $"INSERT INTO build_phase_phase_profile (`build_phase_id`, `phase_profile_id`) VALUES(@Build_Phase_Id, @Phase_Profile_Id)";

                using MySqlCommand command = new MySqlCommand(query, connection);
                command.Parameters.AddWithValue($"@Build_Phase_Id", ids[0]);
                command.Parameters.AddWithValue($"@Phase_Profile_Id", ids[1]);
                command.ExecuteNonQuery();

            }
            Console.WriteLine("Adding BuildPhaseProfile was successful!");

        }


        //build_phase_parameters پارامتر مرحله تولید 
        static void InsertBuildPhaseParametersIntoMySql(int phaseId, string connectionString)
        {
            var parameters = new List<Tuple<string, HashSet<string>>>()
            {
                new("دنیر",disnValue),
                new("فیلامنت", t_sanValue),
                new("رنگ",colorValue),
                new("همبافت",hambaftValue),
                new("تعداد دوک",n_dokValue),
                new("کیفیت",kindValue),
                new("دوک",k_dokValue),
            };


            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();

                var query = $"INSERT INTO build_phase_parameters (`title`, `type`,`build_phase_id`) VALUES(@Title, @Type, @Build_Phase_Id)";

                foreach (var item in parameters)
                {
                    using MySqlCommand command = new MySqlCommand(query, connection);
                    command.Parameters.AddWithValue($"@Title", item.Item1);
                    command.Parameters.AddWithValue($"@Type", "1");
                    command.Parameters.AddWithValue($"@Build_Phase_Id", phaseId);
                    command.ExecuteNonQuery();

                    var getPhasesQuery = "SELECT * FROM build_phase_parameters WHERE title = @Title";
                    using (MySqlCommand getPhsesCommand = new MySqlCommand(getPhasesQuery, connection))
                    {
                        getPhsesCommand.Parameters.AddWithValue("@Title", item.Item1);

                        using (MySqlDataReader reader = getPhsesCommand.ExecuteReader())
                        {
                            reader.Read();
                            var parameterId = reader.GetInt32("id");
                            reader.Close();

                            var insertValueQuery = $"INSERT INTO build_phase_parameter_options (`title`, `build_phase_parameter_id`) VALUES(@Title, @Build_Phase_Parameter_Id)";
                            foreach (var value in item.Item2)
                            {
                                using (MySqlCommand insertProductCommand = new MySqlCommand(insertValueQuery, connection))
                                {
                                    insertProductCommand.Parameters.AddWithValue("@Title", value);
                                    insertProductCommand.Parameters.AddWithValue("@Build_Phase_Parameter_Id", parameterId);
                                    insertProductCommand.ExecuteNonQuery();
                                }
                            }
                          

                        }
                    }
                }
                connection.Close();
            }
            Console.WriteLine("Adding BuildPhaseParameters was successful!");

        }




        //units واحد
        static void InsertUnitIntoMySql(HashSet<string> data, string connectionString)
        {
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();

                foreach (var row in data)
                {
                    var query = $"INSERT INTO units (`title`, `type`) VALUES(@Title, @Type)";

                    using MySqlCommand command = new MySqlCommand(query, connection);
                    command.Parameters.AddWithValue($"@Title", row);
                    command.Parameters.AddWithValue($"@Type", 1);
                    command.ExecuteNonQuery();
                }
            }
            Console.WriteLine("Adding Unit was successful!");
        }


        //products کالا
        static void InsertProductIntoMySql(List<product> data, int profileId, string connectionString)
        {
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                var getUnitQuery = "SELECT * FROM units WHERE title = @Title";
                var getProduct = "SELECT * FROM products WHERE code = @Code";
                var getParameter = "SELECT * FROM build_phase_parameter_options WHERE title in(@Title1, @Title2, @Title3, @Title4)";
                var insertProductQuery = "INSERT INTO products (title, code, unit_id, phase_profile_id) VALUES (@Title, @Code, @Unit_Id, @Phase_Profile_Id)";
                var attachProductWithBuildPhaseParameterOptions =
                    "INSERT INTO build_phase_parameter_option_product (`bppo_id`, `product_id`) VALUES (@OptionId, @ProductId)";

                connection.Open();

                foreach (var row in data)
                {
                    using (MySqlCommand getUnitCommand = new MySqlCommand(getUnitQuery, connection))
                    {
                        getUnitCommand.Parameters.AddWithValue("@Title", row.vahed);

                        using (MySqlDataReader reader = getUnitCommand.ExecuteReader())
                        {
                            reader.Read();
                            var unitId = reader.GetInt32("id");
                            reader.Close();
                            var buildPhaseParameterOptions = new List<int>();
                            using (MySqlCommand getBuildPhaseParameterOptions = new MySqlCommand(getParameter, connection))
                            {
                                getBuildPhaseParameterOptions.Parameters.AddWithValue("@Title1", row.color);
                                getBuildPhaseParameterOptions.Parameters.AddWithValue("@Title2", row.kind);
                                getBuildPhaseParameterOptions.Parameters.AddWithValue("@Title3", row.t_san);
                                getBuildPhaseParameterOptions.Parameters.AddWithValue("@Title4", row.disn);
                                using (MySqlDataReader buildPhaseParameterOptionsReader =
                                       getBuildPhaseParameterOptions.ExecuteReader())
                                {
                                    while (buildPhaseParameterOptionsReader.Read())
                                    {
                                        buildPhaseParameterOptions.Add(buildPhaseParameterOptionsReader.GetInt32("id"));
                                    }
                                    
                                    buildPhaseParameterOptionsReader.Close();
                                }

                            }
                            using (MySqlCommand insertProductCommand = new MySqlCommand(insertProductQuery, connection))
                            {
                                // Use distinct parameter names for each value
                                insertProductCommand.Parameters.AddWithValue("@Title", row.name_kala);
                                insertProductCommand.Parameters.AddWithValue("@Code", row.code_kala); // Replace with the actual code column
                                insertProductCommand.Parameters.AddWithValue("@Unit_Id", unitId);
                                insertProductCommand.Parameters.AddWithValue("@Phase_Profile_Id", profileId);

                                insertProductCommand.ExecuteNonQuery();
                            }

                            var productId = 0;

                            using (MySqlCommand getProductCommand = new MySqlCommand(getProduct, connection))
                            {
                                getProductCommand.Parameters.AddWithValue("@Code", row.code_kala);

                                using (MySqlDataReader productReader = getProductCommand.ExecuteReader())
                                {
                                    productReader.Read();
                                    productId = productReader.GetInt32("id");
                                    productReader.Close();
                                }
                            }

                            foreach (var itemId in buildPhaseParameterOptions)
                            {
                                using (MySqlCommand attachProductWithBuildPhaseParameterOptionsCommand =
                                       new MySqlCommand(attachProductWithBuildPhaseParameterOptions, connection))
                                {

                                    attachProductWithBuildPhaseParameterOptionsCommand.Parameters.AddWithValue(
                                        "@OptionId", itemId);
                                    attachProductWithBuildPhaseParameterOptionsCommand.Parameters.AddWithValue(
                                        "@ProductId", productId);
                                    attachProductWithBuildPhaseParameterOptionsCommand.ExecuteNonQuery();


                                }
                            }
                        }

                    }

                }
            }
            Console.WriteLine("Adding Product was successful!");

        }



        //out_products رسید تولید
        static void InsertOutProductIntoMySql(List<outProduct> data, string connectionString)
        {
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                var getProductQuery = "SELECT * FROM products WHERE code = @Code";
                var insertOutProductQuery = "INSERT INTO out_products (product_id, barcode, value, value_nakhales) VALUES (@Product_Id, @Barcode, @Value, @Value_Nakhales)";

                connection.Open();

                foreach (var row in data)
                {
                    using (MySqlCommand getUnitCommand = new MySqlCommand(getProductQuery, connection))
                    {
                        getUnitCommand.Parameters.AddWithValue("@Code", row.code_kala);

                        using (MySqlDataReader reader = getUnitCommand.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                var productId = reader.GetInt32("id");
                                reader.Close();
                                using MySqlCommand insertProductCommand = new MySqlCommand(insertOutProductQuery, connection);
                                // Use distinct parameter names for each value
                                insertProductCommand.Parameters.AddWithValue("@Product_Id", productId);
                                insertProductCommand.Parameters.AddWithValue("@Barcode", row.no_c);
                                insertProductCommand.Parameters.AddWithValue("@Value", row.wnet);
                                insertProductCommand.Parameters.AddWithValue("@Value_Nakhales", row.nwnet); // Replace with the actual code column

                                insertProductCommand.ExecuteNonQuery();
                            }

                        }

                    }

                }
            }
            Console.WriteLine("Adding OutProduct was successful!");

        }




    }

}

