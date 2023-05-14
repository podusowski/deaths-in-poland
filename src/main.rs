use std::collections::HashMap;

use calamine::{open_workbook, Error, Reader, Xlsx};
use plotters::prelude::*;
use prettytable::{format, Cell, Row};

#[derive(Debug, Clone)]
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

#[derive(Clone)]
struct AnnualData {
    year: usize,
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

const AGE_GROUPS: &'static [&str] = &[
    "0 - 4",
    "5 - 9",
    "10 - 14",
    "15 - 19",
    "20 - 24",
    "25 - 29",
    "30 - 34",
    "35 - 39",
    "40 - 44",
    "45 - 49",
    "50 - 54",
    "55 - 59",
    "60 - 64",
    "65 - 69",
    "70 - 74",
    "75 - 79",
    "80 - 84",
    "85 - 89",
    "90 i więcej",
];

fn find_sheet(year: usize) -> std::path::PathBuf {
    for p in [
        std::path::Path::new(format!("data/Zgony wedêug tygodni w Polsce_{}.xlsx", year).as_str()),
        std::path::Path::new(format!("data/Zgony wedlug tygodni w Polsce_{}.xlsx", year).as_str()),
        std::path::Path::new(format!("data/Zgony wedИug tygodni w Polsce_{}.xlsx", year).as_str()),
        std::path::Path::new(format!("data/Zgony wedug tygodni w Polsce_{}.xlsx", year).as_str()),
    ] {
        if p.exists() {
            return p.to_owned();
        }
    }
    panic!("no path");
}

fn read(year: usize) -> anyhow::Result<AnnualData> {
    let path = find_sheet(year);
    let mut workbook: Xlsx<_> = open_workbook(&path)?;

    let range = workbook
        .worksheet_range("OGÓŁEM")
        .ok_or(Error::Msg("No sheet."))??;

    let mut age_groups = HashMap::<&str, AgeGroup>::new();
    for age_group in AGE_GROUPS {
        age_groups.insert(&age_group, find_age_group(&range, age_group)?);
    }

    Ok(AnnualData {
        year,
        title: format!("{}", path.display()),
        general: find_age_group(&range, "Ogółem")?,
        age_groups,
    })
}

fn flatten_out_into_weeks(years: &[AnnualData]) -> Vec<Vec<u32>> {
    let mut data = Vec::<Vec<u32>>::new();
    for age_group in AGE_GROUPS {
        let mut deaths_in_age_group = Vec::<u32>::new();
        for year in years {
            deaths_in_age_group.extend(year.age_groups[age_group].0.iter());
        }
        data.push(deaths_in_age_group);
    }
    data
}

fn draw_super_plot(years: &[AnnualData]) -> anyhow::Result<()> {
    let area = BitMapBackend::gif("output/super.gif", (1024, 760), 100)?.into_drawing_area();

    let data = flatten_out_into_weeks(years);
    let x_axis = 0u32..data.len() as u32;
    let z_axis = 0u32..data[0].len() as u32; // They all should have the same length.

    for yaw in 0..360 {
        area.fill(&BLACK)?;
        let mut chart = ChartBuilder::on(&area).build_cartesian_3d(
            x_axis.clone(),
            0u32..3000u32,
            z_axis.clone(),
        )?;

        chart.with_projection(|mut pb| {
            pb.pitch = 0.7;
            pb.yaw = yaw as f64 * (3.14 / 180.0);
            pb.scale = 0.7;
            pb.into_matrix()
        });

        chart
            .configure_axes()
            .label_style(TextStyle::from(("sans-serif", 14)).color(&RED))
            .x_formatter(&|x| AGE_GROUPS[*x as usize].to_string())
            .z_formatter(&|_| "".to_string())
            .draw()?;

        chart.draw_series(
            SurfaceSeries::xoz(x_axis.clone(), z_axis.clone(), |group, week| {
                data[group as usize][week as usize]
            })
            .style_func(&|&v| RED.mix(v as f64 / 6000.0 + 0.2).into()),
        )?;

        area.present()?;
    }

    Ok(())
}

fn draw_plot_for_age_group(years: &[AnnualData], age_group: &str) -> anyhow::Result<()> {
    let path = format!("output/age-group-{}.svg", age_group);
    let area = SVGBackend::new(path.as_str(), (800, 400)).into_drawing_area();

    let x_axis = 0u32..years[0].age_groups[age_group].0.len() as u32; // They all should have the same length.

    let min = years
        .iter()
        .map(|year| year.age_groups[age_group].0.iter().min().unwrap_or(&0))
        .min()
        .unwrap_or(&0);

    let max = years
        .iter()
        .map(|year| year.age_groups[age_group].0.iter().max().unwrap_or(&0))
        .max()
        .unwrap_or(&0);

    let y_axis = *min..*max;

    let start_year = &years[0].year;
    let end_year = &years[years.len() - 1].year;

    let caption = format!(
        "Zgony w latach {} - {} wśród osób {}",
        start_year, end_year, age_group
    );

    area.fill(&WHITE)?;

    let mut years = years.to_vec();
    years.sort_by(|a, b| a.year.cmp(&b.year));
    years.reverse();

    let shades_of_gray = (1..5).map(|n| BLACK.mix((5.0 - n as f64) / 5.0));

    let colors = [CYAN, MAGENTA]
        .iter()
        .map(|c| c.to_rgba())
        .chain(shades_of_gray);

    let mut chart = ChartBuilder::on(&area)
        .caption(
            caption.clone(),
            ("sans-serif", 12).into_font().color(&BLACK),
        )
        .set_label_area_size(LabelAreaPosition::Left, 12.percent())
        .set_label_area_size(LabelAreaPosition::Bottom, 10.percent())
        .margin(1.percent())
        .build_cartesian_2d(x_axis.clone(), y_axis.clone())?;

    chart
        .configure_mesh()
        .disable_mesh()
        .x_desc("Tydzień")
        .y_desc("Ilość zgonów")
        .draw()?;

    for (year, color) in years.iter().zip(colors) {
        chart
            .draw_series(LineSeries::new(
                year.age_groups[age_group]
                    .0
                    .iter()
                    .enumerate()
                    .map(|(x, y)| (x as u32, *y)),
                color.stroke_width(2),
            ))?
            .label(format!("{}", year.year))
            .legend(move |(x, y)| Rectangle::new([(x, y - 5), (x + 10, y + 5)], color.filled()));
    }

    chart
        .configure_series_labels()
        .position(SeriesLabelPosition::UpperMiddle)
        .border_style(&BLACK)
        .draw()?;

    area.present()?;

    Ok(())
}

fn draw_annual_sums(years: &[AnnualData]) -> anyhow::Result<()> {
    let path = "output/annual.png";
    let area = BitMapBackend::new(path, (1024, 760)).into_drawing_area();

    let x_axis = 0u32..years[0].general.0.len() as u32; // They all should have the same length.
    let y_axis = 0u32..20000u32;

    area.fill(&WHITE)?;

    let colors = years
        .iter()
        .enumerate()
        .map(|(number, _)| RED.mix(1.0 - (number as f64 / 5.0)));

    for (year, color) in years.iter().zip(colors) {
        let mut chart =
            ChartBuilder::on(&area).build_cartesian_2d(x_axis.clone(), y_axis.clone())?;

        chart
            .configure_mesh()
            .disable_x_mesh()
            .disable_y_mesh()
            .draw()?;

        chart.draw_series(LineSeries::new(
            year.general
                .0
                .iter()
                .enumerate()
                .map(|(x, y)| (x as u32, *y)),
            color,
        ))?;

        area.present()?;
    }

    Ok(())
}

fn print_tables(years: &[AnnualData]) {
    let mut table = prettytable::Table::new();
    let github = format::FormatBuilder::new()
        .column_separator('|')
        .borders('|')
        .separators(
            &[format::LinePosition::Intern],
            format::LineSeparator::new('-', '|', '|', '|'),
        )
        .padding(1, 1)
        .build();
    table.set_format(github);

    table.add_row(Row::new({
        let mut row = vec![Cell::new("")];
        row.extend(
            years
                .iter()
                .map(|year| Cell::new(format!("{}", year.year).as_str())),
        );
        row
    }));

    table.add_row(Row::new({
        let mut row = vec![Cell::new("średnia tygodniowa")];
        row.extend(
            years
                .iter()
                .map(|year| Cell::new(format!("{}", year.general.avg().round()).as_str())),
        );
        row
    }));

    for age_group in AGE_GROUPS {
        table.add_row(Row::new({
            let mut row = vec![Cell::new(age_group)];
            row.extend(years.iter().map(|year| {
                Cell::new(format!("{}", year.age_groups[age_group].avg().round()).as_str())
            }));
            row
        }));
    }

    table.printstd();
}

fn main() -> anyhow::Result<()> {
    const OUTPUT_DIR: &str = "output";
    //std::fs::create_dir(OUTPUT_DIR)
    //    .expect(format!("Can't create directory '{}'", OUTPUT_DIR).as_str());
    println!("Result will be written to {}", OUTPUT_DIR);

    let years = (2015..2024)
        .map(|year| read(year).expect("Could not read"))
        .collect::<Vec<_>>();

    for year in &years {
        println!("{}, średnia: {}", year.year, year.general.avg());
        for (label, age_group) in &year.age_groups {
            println!("  {} średnia: {}", label, age_group.avg());
        }
        println!("");
    }

    print_tables(&years);

    draw_super_plot(&years)?;

    for age_group in AGE_GROUPS {
        draw_plot_for_age_group(&years, age_group)?;
    }

    draw_annual_sums(&years)?;

    Ok(())
}
