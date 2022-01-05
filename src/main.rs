use std::collections::HashMap;

use calamine::{open_workbook, Error, Reader, Xlsx};
use plotters::prelude::*;

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
    age_groups: HashMap<&'static str, AgeGroup>,
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

const AGE_GROUPS: &'static [&'static str] = ["0 - 4"];

fn read(path: &str) -> anyhow::Result<AnnualData> {
    let mut workbook: Xlsx<_> = open_workbook(path)?;

    let range = workbook
        .worksheet_range("OGÓŁEM")
        .ok_or(Error::Msg("No sheet."))??;

    let mut age_groups = HashMap::<&str, AgeGroup>::new();
    age_groups.insert("0 - 4", find_age_group(&range, "0 - 4")?);
    age_groups.insert("5 - 9", find_age_group(&range, "5 - 9")?);

    Ok(AnnualData {
        title: path.to_owned(),
        general: find_age_group(&range, "Ogółem")?,
        age_groups,
    })
}

fn draw_plot(years: &[AnnualData]) -> anyhow::Result<()> {
    let area = BitMapBackend::new("plot.png", (1024, 760)).into_drawing_area();
    area.fill(&WHITE)?;
    let x_axis = (2017..2021);
    let z_axis = (1..3);

    let mut chart = ChartBuilder::on(&area)
        .caption(format!("3D Plot Test"), ("sans", 20))
        .build_cartesian_3d(x_axis, -3.0..3.0, z_axis.clone())?;

    chart.with_projection(|mut pb| {
        pb.yaw = 0.5;
        pb.scale = 0.9;
        pb.into_matrix()
    });

    chart.configure_axes().draw()?;

    chart.draw_series(
        SurfaceSeries::xoz(
            [1, 2, 3].iter().cloned(),
            [1, 2, 3].iter().cloned(),
            |year, group| 1.0,
        )
        .style(BLUE.mix(0.2).filled()),
    )?;

    chart
        .configure_series_labels()
        .border_style(&BLACK)
        .draw()?;

    area.present()?;

    Ok(())
}

fn main() -> anyhow::Result<()> {
    let years = [
        read("data/Zgony wedИug tygodni w Polsce_2021.xlsx")?,
        read("data/Zgony wedИug tygodni w Polsce_2020.xlsx")?,
        read("data/Zgony wedИug tygodni w Polsce_2019.xlsx")?,
        read("data/Zgony wedИug tygodni w Polsce_2018.xlsx")?,
        read("data/Zgony wedИug tygodni w Polsce_2017.xlsx")?,
    ];

    for year in &years {
        println!("{:?}", year.title);
        println!("general ({}): {:?}", year.general.avg(), year.general);
        for (label, age_group) in &year.age_groups {
            println!("{} ({}): {:?}", label, age_group.avg(), age_group.0);
        }
        println!("");
    }

    draw_plot(&years)?;

    Ok(())
}
