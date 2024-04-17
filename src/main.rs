use std::io::Read;
use std::{fs::File, net::SocketAddr};

use axum::http::header;
use axum::response::IntoResponse;
use axum::{
    http::{Response, StatusCode},
    routing::get,
    Router,
};

use tokio::net::TcpListener;
use xlsxwriter::{format::*, Workbook};

fn generate_xlsx() -> Result<Vec<u8>, String> {
    let path_file = "target/data.xlsx";

    let workbook = Workbook::new(path_file).expect("msg");

    let mut sheet1 = workbook.add_worksheet(None).expect("msg");
    sheet1
        .write_string(
            0,
            0,
            "Red text",
            Some(
                Format::new()
                    .set_font_size(20.0)
                    .set_bold()
                    .set_border_color(FormatColor::White)
                    .set_border(FormatBorder::Thin)
                    .set_font_color(FormatColor::Red),
            ),
        )
        .expect("msg");
    sheet1
        .write_number(
            0,
            1,
            20.,
            Some(
                Format::new()
                    .set_font_size(20.0)
                    .set_bold()
                    .set_border_color(FormatColor::White)
                    .set_border(FormatBorder::Thin)
                    .set_font_color(FormatColor::Lime),
            ),
        )
        .expect("msg");

    sheet1
        .write_formula_num(
            1,
            0,
            "=4540+B1",
            Some(
                Format::new()
                    .set_bold()
                    .set_border_color(FormatColor::White)
                    .set_border(FormatBorder::Thin)
                    .set_font_color(FormatColor::Orange),
            ),
            0.,
        )
        .expect("msg");
    sheet1
        .write_url(
            1,
            1,
            "https://github.com/informationsea/xlsxwriter-rs",
            Some(
                Format::new()
                    .set_border_color(FormatColor::White)
                    .set_border(FormatBorder::Thin)
                    .set_font_color(FormatColor::Blue)
                    .set_underline(FormatUnderline::Single),
            ),
        )
        .expect("msg");

    sheet1
        .write_formula_num(
            2,
            0,
            "=4540+A2",
            Some(
                Format::new()
                    .set_bold()
                    .set_border_color(FormatColor::White)
                    .set_border(FormatBorder::Thin)
                    .set_font_color(FormatColor::Brown),
            ),
            0.,
        )
        .expect("msg");
    sheet1
        .write_string(
            2,
            1,
            "Red text",
            Some(
                Format::new()
                    .set_font_color(FormatColor::Red)
                    .set_border_color(FormatColor::White)
                    .set_border(FormatBorder::Thin),
            ),
        )
        .expect("msg");

    sheet1
        .write_formula_num(
            3,
            0,
            "=4540+A3",
            Some(
                Format::new()
                    .set_bold()
                    .set_font_color(FormatColor::Blue)
                    .set_border_color(FormatColor::White)
                    .set_border(FormatBorder::Thin),
            ),
            0.,
        )
        .expect("msg");
    sheet1
        .write_string(
            3,
            1,
            "Red text",
            Some(
                Format::new()
                    .set_font_color(FormatColor::Red)
                    .set_border_color(FormatColor::White)
                    .set_border(FormatBorder::Thin),
            ),
        )
        .expect("msg");

    sheet1.set_selection(1, 0, 1, 2);
    sheet1.set_tab_color(FormatColor::Cyan);
    let _ = workbook.close();

    let mut f = File::open(path_file).expect("msg");
    let mut data = vec![];
    let _ = f
        .read_to_end(&mut data)
        .map_err(|e| e.to_string())
        .expect("");

    Ok(data)
}

async fn excel_handler() -> impl IntoResponse {
    match generate_xlsx() {
        Ok(buffer) => (
            [
                (
                    header::CONTENT_TYPE,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ),
                (
                    header::CONTENT_DISPOSITION,
                    "attachment; filename=data.xlsx",
                ),
            ],
            buffer,
        ),
        Err(_) => (
            [
                (
                    header::CONTENT_TYPE,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ),
                (
                    header::CONTENT_DISPOSITION,
                    "attachment; filename=error.xlsx",
                ),
            ],
            Vec::new(),
        ),
    }
}

#[tokio::main]
async fn main() -> Result<(), Box<dyn std::error::Error>> {
    let _ = dotenv::dotenv();
    let port = std::env::var("PORT").expect("PORT must be set");

    let app = Router::new().route("/", get(excel_handler));

    let addr = SocketAddr::from(([0, 0, 0, 0], port.parse::<u16>().unwrap()));
    let listener = TcpListener::bind(addr).await?;

    axum::serve(listener, app).await.unwrap();

    Ok(())
}
