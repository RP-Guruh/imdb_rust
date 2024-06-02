use rust_xlsxwriter::*;

fn main() {
    let url = "https://www.imdb.com/search/title/?groups=top_100&sort=user_rating,desc&count=100";
    println!("Hi, scrap from : {}", url);

    let response = reqwest::blocking::get(
        url,
    )
        .unwrap()
        .text()
        .unwrap();

    let document = scraper::Html::parse_document(&response);

    let title_selector = scraper::Selector::parse("h3.ipc-title__text").unwrap();
    let rate_selector = scraper::Selector::parse("span.ipc-rating-star").unwrap();

    //let titles = document.select(&title_selector).map(|x| x.inner_html());
    let titles: Vec<String> = document
        .select(&title_selector)
        .map(|x| x.inner_html())
        .collect();

    let ratings: Vec<String> = document
        .select(&rate_selector)
        .filter_map(|element| element.value().attr("aria-label"))
        .map(|attr| attr.to_string())
        .collect();

    let _ = save_excell(titles, ratings);
}

fn save_excell(titles: Vec<String>, rate: Vec<String>) -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let bold_format = Format::new().set_bold();

    let worksheet = workbook.add_worksheet();

    // Set the column width for clarity.
    worksheet.set_column_width(0, 20)?;

    // Write a string without formatting.
    worksheet.write_with_format(0, 0, "No.", &bold_format)?;
    worksheet.write_with_format(0, 1, "Film", &bold_format)?;
    worksheet.write_with_format(0, 2, "Rating", &bold_format)?;
    let mut i = 1;
    for title in titles {
        if i > 100 {
            i = 1;
            break;
        }
        worksheet.write(i, 1, title.splitn(2, '.').nth(1).unwrap_or("").trim())?;
        worksheet.write(i, 0, i)?;
        i += 1;
    }

    for imdb in rate {
        if i > 100 {
            break;
        }
        worksheet.write(i, 2, imdb.splitn(12, ':').nth(1).unwrap_or("").trim())?;
        i += 1;
    }

    // Save the file to disk.
    workbook.save("demo.xlsx")?;
    println!("Finish");
    Ok(())
}