use sqlx::{Mssql, Pool};
use tokio;
use xlsxwriter::*;

#[tokio::main]
async fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Database connection string
    let db_url = "mssql://ccmadmin:ccm@486*@172.16.8.211:1433/ccms10_New";

    // Create a connection pool
    let pool = Pool::<Mssql>::connect(db_url).await?;
    
    // SQL Query
    let rows = sqlx::query!("SELECT top 10 masterkey FROM OMSJobQueueItems with (NoLock)")
        .fetch_all(&pool)
        .await?;
    
   
    let file_path = "D:/Devilal/employees.xlsx";

    // Create an Excel workbook & worksheet
    let workbook = Workbook::new(file_path);
    let mut sheet = workbook.add_worksheet(None)?;

    // Write header row
    sheet.write_string(0, 0, "MasterKey", None)?;


    // Write data rows
    for (i, row) in rows.iter().enumerate() {
        sheet.write_string(i as u32 + 1, 1, &row.masterkey, None)?;
    }

    // Save and close the workbook
    workbook.close()?;
    
    println!("✅ Data exported to: {}", file_path);
    Ok(())
}
