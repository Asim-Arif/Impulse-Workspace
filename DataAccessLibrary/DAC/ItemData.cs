using Dapper;
using DataAccessLibrary.Models;
using Microsoft.Extensions.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;


namespace DataAccessLibrary
{
    public class ItemData : IItemData
    {
        private readonly IConfiguration _config;

        public ItemData(IConfiguration config)
        {
            _config = config;
        }

        private string GetConnectionString()
        {
            return _config.GetConnectionString("DefaultConnection");
        }

        public async Task InsertItemAsync(Item item)
        {
            using (IDbConnection db = new SqlConnection(GetConnectionString()))
            {
                // SQL query to insert an item into the database
                string sql = @"INSERT INTO Items (ItemID,ItemName, Unit, Opening_Stock, ItemPrice,ItemSize,CatID) 
                                OUTPUT INSERTED.EntryID
                               VALUES (@ItemID,@ItemName, @Unit, @Quantity, @Price,'1','02')";
                // Execute the query using Dapper
                int iEntryID = await db.QuerySingleAsync<int>(sql, item);
                //await db.ExecuteAsync(sql, item);

                foreach (var rmDetail in item.RMDetails) 
                {
                    sql = @"INSERT INTO Items_RM(ItemID,RMCode,RMQty,Items_RefID) VALUES(@ItemID,@RMCode,@RMQty,@Items_RefID)";
                    var rmParams = new {ItemID = item.ItemID, RMCode = rmDetail.RMCode, RMQty = rmDetail.RMQty, Items_RefID = iEntryID };
                    await db.ExecuteAsync(sql, rmParams);
                }

            }
        }

        public async Task<Item> GetItemByIdAsync(string itemid)
        {
            using (IDbConnection db = new SqlConnection(GetConnectionString()))
            {
                string sql = "SELECT ItemID,ItemName,Unit,Opening_Stock AS Quantity,ItemPrice AS Price FROM Items WHERE ItemID = '@Id'";
                return await db.QueryFirstOrDefaultAsync<Item>(sql, new { Id = itemid });
            }
        }

        public async Task<IEnumerable<Item>> GetItemsAsync()
        {
            using (IDbConnection db = new SqlConnection(GetConnectionString()))
            {
                string sql = "SELECT ItemID,ItemName,Unit,Opening_Stock AS Quantity,ItemPrice AS Price,Catagory FROM VItems";
                return await db.QueryAsync<Item>(sql);
            }
        }
        public async Task<List<Material>> GetMaterialsAsync()
        {
            using (IDbConnection db = new SqlConnection(GetConnectionString()))
            {
                //string sql = "SELECT MaterialID,MaterialName FROM Material";
                string sql = "SELECT ItemID AS MaterialID,ItemName AS MaterialName,Unit,InHand AS StockQty FROM VItems";
                return (await db.QueryAsync<Material>(sql)).ToList();
            }
        }

        public async Task<int> GetmaxBookingNumberAsync()
        {
            using (IDbConnection db = new SqlConnection(GetConnectionString()))
            {
                var maxNumber = await db.ExecuteScalarAsync<int>(
                    "SELECT MAX(CAST(RIGHT(InvoiceNo,LEN(InvoiceNo)-9) AS INT)) FROM ItemSales_Booking WHERE LEFT(InvoiceNo,3)='SBI'");
                return maxNumber;
            }
        }

        public async Task InsertSaleBookingAsync(Item item)
        {
            Console.WriteLine("came into save function...");
            //if (item == null)
            //{
            //    throw new ArgumentNullException(nameof(item), "Item cannot be null.");
            //}

            //if (item.RMDetails == null || !item.RMDetails.Any())
            //{
            //    throw new ArgumentException("Item must have at least one RMDetail.", nameof(item));
            //}

            using (IDbConnection db = new SqlConnection(GetConnectionString()))
            {
                try
                {
                    Console.WriteLine("Attempting to open database connection...");
                    db.Open(); // Ensure the connection is opened

                    // SQL query to insert Sales Booking into the database
                    string sql = @"INSERT INTO ItemSales_Booking (InvoiceNo, DT, CustCode, Country, Discount, ServiceCharges, SaleDescription) 
                           OUTPUT INSERTED.EntryID
                           VALUES (@BookInvoiceNo, @BookDT, @CustCode, 'Pakistan', 0, 0, @BookDescription)";

                    // Execute the query using Dapper and get the generated EntryID
                    int iEntryID = await db.QuerySingleAsync<int>(sql, item);
                    Console.WriteLine($"Booking inserted with EntryID: {iEntryID}");

                    // Insert each RMDetail in the ItemSales_Items_Booking table
                    foreach (var rmDetail in item.RMDetails)
                    {
                        if (rmDetail == null)
                        {
                            Console.WriteLine("RMDetail is null, skipping.");
                            continue; // Skip null RMDetails
                        }

                        sql = @"INSERT INTO ItemSales_Items_Booking (IS_RefID, ItemID, Qty, Rate, GQty) 
                        VALUES (@Items_RefID, @ItemID, @Qty, 0, @GQty)";

                        // Use values from each rmDetail for the parameters
                        var rmParams = new
                        {
                            Items_RefID = iEntryID,
                            ItemID = rmDetail.ItemID,
                            Qty = rmDetail.Qty,
                            GQty = rmDetail.GQty
                        };

                        await db.ExecuteAsync(sql, rmParams);
                        Console.WriteLine($"Inserted RMDetail for ItemID: {rmDetail.ItemID} with Qty: {rmDetail.Qty}");
                    }

                    Console.WriteLine("All RMDetails inserted successfully.");
                }
                catch (SqlException sqlEx)
                {
                    Console.WriteLine($"SQL Error: {sqlEx.Message}");
                    throw; // Optionally rethrow the exception
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occurred: {ex.Message}");
                    throw; // Optionally rethrow the exception
                }
                finally
                {
                    // Optional: Clean up if needed
                    db.Close(); // Ensure the database connection is closed
                }
            }
        }

        public async Task<List<Item>> GetSalesBookingListAsync()
        {
            using (IDbConnection db = new SqlConnection(GetConnectionString()))
            {
                string sql = "SELECT TOP 50 InvoiceNo AS SBInvoiceNo, DT AS SBDT, CustCode AS Customer, TotalAmount AS TotalQty, SoldQty As SoldQty, BalQty AS BalQty FROM VItemSales_Booking ORDER BY VItemSales_Booking.DT DESC";
                // Fetch SBList data
                List<SBList> sbList = (await db.QueryAsync<SBList>(sql)).ToList();
                return ConvertSBListToItemList(sbList);
            }
        }
        private List<Item> ConvertSBListToItemList(List<SBList> sbList)
        {
            return sbList.Select(sb =>
            {
                //Console.WriteLine($"Reading Data: {sb.SBInvoiceNo}, {sb.Customer}, {sb.SBDT}, {sb.SoldQty}, {sb.BalQty}");

                return new Item
                {
                    SBInvoiceNo = sb.SBInvoiceNo,   // Adjust property names based on Item class
                    SBDT = sb.SBDT,
                    Customer = sb.Customer,
                    TotalQty = sb.TotalQty,
                    SoldQty = sb.SoldQty,
                    BalQty = sb.BalQty,
                    // Map other properties as needed
                };
            }).ToList();
        }
    }
}
