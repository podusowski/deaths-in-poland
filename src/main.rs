use calamine::{open_workbook, Error, RangeDeserializerBuilder, Reader, Xlsx};

fn main() -> anyhow::Result<()> {
    let mut workbook: Xlsx<_> =
        open_workbook("data/Zgony wedИug tygodni w Polsce_2021.xlsx").unwrap();

    let range = workbook
        .worksheet_range("OGÓŁEM")
        .ok_or(Error::Msg("No sheet."))??;

    //for row in range.rows() {
    //    println!("{:?}", row);
    //}

    let general = &range
        .rows()
        .find(|r| {
            if let Some(s) = r[0].get_string() {
                s == "Ogółem"
            } else {
                false
            }
        })
        .ok_or(anyhow::anyhow!("no data"))?[3..];

    println!("{:?}", general);

    Ok(())

    //let mut iter = RangeDeserializerBuilder::new().from_range(&range)?;

    //if let Some(result) = iter.next() {
    //    let (label, value): (String, f64) = result?;
    //    assert_eq!(label, "celsius");
    //    assert_eq!(value, 22.2222);
    //    Ok(())
    //} else {
    //    Err(anyhow::anyhow!("expected at least one record but got none"))
    //}
}
