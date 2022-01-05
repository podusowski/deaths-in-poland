use calamine::{open_workbook, Error, RangeDeserializerBuilder, Reader, Xlsx};

struct AnnualData {
    general: Vec<Option<f64>>,
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
        .map(|v| v.get_float().or(None))
        .collect();

    Ok(AnnualData { general })
}

fn main() -> anyhow::Result<()> {
    let y2021 = read("data/Zgony wedИug tygodni w Polsce_2021.xlsx")?;
    println!("{:?}", y2021.general);
    Ok(())
}
