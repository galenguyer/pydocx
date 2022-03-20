use docx_rs::*;

use std::fs::File;
use std::io::Read;

pub fn main() {
    let mut file = File::open("./python.docx").unwrap();
    let mut buf = vec![];
    file.read_to_end(&mut buf).unwrap();

    // let mut file = File::create("./python.json").unwrap();
    let res = read_docx(&buf).unwrap().document;
    for document_child in res.children {
        match document_child {
            DocumentChild::Paragraph(paragraph) => {
                for paragraph_child in paragraph.children {
                    match paragraph_child {
                        ParagraphChild::Run(run) => {
                            for run_child in run.children {
                                match run_child {
                                    RunChild::Tab(_) => {print!("    ")}
                                    RunChild::Text(text) => {println!("{}", text.text)}
                                    _ => {}
                                }
                            }
                        },
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }
    // file.write_all(res.as_bytes()).unwrap();
    // file.flush().unwrap();
}
