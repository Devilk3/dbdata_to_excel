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
    
    // Path for the Excel file
    let file_path = "D:/Devilal/employees.xlsx";

    // Create an Excel workbook & worksheet
    let workbook = Workbook::new(file_path);
    let mut sheet = workbook.add_worksheet(None)?;

    // Write header row
    sheet.write_string(0, 0, "MasterKey", None)?;
   // sheet.write_string(0, 1, "Name", None)?;
   // sheet.write_string(0, 2, "Age", None)?;

    // Write data rows
    for (i, row) in rows.iter().enumerate() {
        sheet.write_number(i as u32 + 1, 0, row.id as f64, None)?;
        sheet.write_string(i as u32 + 1, 1, &row.name, None)?;
        sheet.write_number(i as u32 + 1, 2, row.age as f64, None)?;
    }

    // Save and close the workbook
    workbook.close()?;
    
    println!("✅ Data exported to: {}", file_path);
    Ok(())
}
