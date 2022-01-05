use calamine::{open_workbook, Error, Reader, Xlsx};

#[derive(Debug)]
struct AgeGroup(Vec<u32>);

impl AgeGroup {
    /// Calculate an average out of non-zero elements.
    fn avg(&self) -> f32 {
        // Ugh, need this imperative code to count it in a single pass.
        let mut sum = 0;
        let mut count = 0;
        for e in self.0.iter().filter(|v| **v != 0) {
            sum += e;
            count += 1;
        }
        sum as f32 / count as f32
    }
}

struct AnnualData {
    title: String,
    general: AgeGroup,
    _0to4: AgeGroup,
    _5to9: AgeGroup,
    _65to69: AgeGroup,
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
        title: path.to_owned(),
        general: find_age_group(&range, "Ogółem")?,
        _0to4: find_age_group(&range, "0 - 4")?,
        _5to9: find_age_group(&range, "5 - 9")?,
        _65to69: find_age_group(&range, "65 - 69")?,
    })
}

fn main() -> anyhow::Result<()> {
    let years = [
        read("data/Zgony wedИug tygodni w Polsce_2021.xlsx")?,
        read("data/Zgony wedИug tygodni w Polsce_2020.xlsx")?,
        read("data/Zgony wedИug tygodni w Polsce_2019.xlsx")?,
        read("data/Zgony wedИug tygodni w Polsce_2018.xlsx")?,
        read("data/Zgony wedИug tygodni w Polsce_2017.xlsx")?,
    ];

    for year in years {
        println!("{:?}", year.title);
        println!("general ({}): {:?}", year.general.avg(), year.general);
        println!("0-4 ({}): {:?}", year._0to4.avg(), year._0to4);
        println!("5-9 ({}): {:?}", year._5to9.avg(), year._5to9);
        println!("65-69 ({}): {:?}", year._65to69.avg(), year._65to69);
        println!("");
    }

    Ok(())
}
