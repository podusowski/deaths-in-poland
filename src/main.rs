use calamine::{open_workbook, Error, RangeDeserializerBuilder, Reader, Xlsx};

struct AnnualData {
    general: Vec<u32>,
}

fn read(path: &str) -> anyhow::Result<AnnualData> {
    let mut workbook: Xlsx<_> = open_workbook(path)?;

    let range = workbook
        .worksheet_range("OGÓŁEM")
        .ok_or(Error::Msg("No sheet."))??;

    let general = range
        .rows()
        .find(|r| {
            if let Some(s) = r[0].get_string() {
                s == "Ogółem"
            } else {
                false
            }
        })
        .ok_or(anyhow::anyhow!("no data"))?[3..]
        .iter()
        .map(|v| v.get_float().unwrap_or(0.0) as u32)
        .collect();

    Ok(AnnualData { general })
}

fn read_and_print(path: &str) -> anyhow::Result<()> {
    println!("{:?}", path);
    println!("{:?}", read(path)?.general);
    println!("");
    Ok(())
}

fn main() -> anyhow::Result<()> {
    read_and_print("data/Zgony wedИug tygodni w Polsce_2021.xlsx")?;
    read_and_print("data/Zgony wedИug tygodni w Polsce_2020.xlsx")?;
    read_and_print("data/Zgony wedИug tygodni w Polsce_2019.xlsx")?;
    read_and_print("data/Zgony wedИug tygodni w Polsce_2018.xlsx")?;
    read_and_print("data/Zgony wedИug tygodni w Polsce_2017.xlsx")?;
    Ok(())
}
