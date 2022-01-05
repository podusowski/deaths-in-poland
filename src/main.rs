use calamine::{open_workbook, Error, Reader, Xlsx};

#[derive(Debug)]
struct AgeGroup(Vec<u32>);

impl AgeGroup {
    fn avg(&self) -> f32 {
        self.0.iter().sum::<u32>() as f32 / self.0.len() as f32
    }
}

struct AnnualData {
    general: AgeGroup,
    _0to4: AgeGroup,
    _5to9: AgeGroup,
}

fn find_age_group(
    range: &calamine::Range<calamine::DataType>,
    group: &str,
) -> anyhow::Result<AgeGroup> {
    Ok(AgeGroup(
        range
            .rows()
            // Find the right group and take the data for the whole country.
            .find(|row| row[0].get_string() == Some(group) && row[1].get_string() == Some("PL"))
            // Drop labels and leave only the numbers.
            .ok_or(anyhow::anyhow!("no data"))?[3..]
            .iter()
            // Calamine reads it as floats, but we want decimals instead.
            .map(|v| v.get_float().unwrap_or(0.0) as u32)
            .collect(),
    ))
}

fn read(path: &str) -> anyhow::Result<AnnualData> {
    let mut workbook: Xlsx<_> = open_workbook(path)?;

    let range = workbook
        .worksheet_range("OGÓŁEM")
        .ok_or(Error::Msg("No sheet."))??;

    Ok(AnnualData {
        general: find_age_group(&range, "Ogółem")?,
        _0to4: find_age_group(&range, "0 - 4")?,
        _5to9: find_age_group(&range, "5 - 9")?,
    })
}

fn read_and_print(path: &str) -> anyhow::Result<()> {
    let annual = read(path)?;
    println!("{:?}", path);
    println!("general ({}): {:?}", annual.general.avg(), annual.general);
    println!("0-4 ({}): {:?}", annual._0to4.avg(), annual._0to4);
    println!("5-9 ({}): {:?}", annual._5to9.avg(), annual._5to9);
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
