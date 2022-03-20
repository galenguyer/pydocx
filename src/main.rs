use docx_rs::*;

use std::env;
use std::fs::File;
use std::io::Read;

pub fn main() {
    let file_name = env::args().nth(1).unwrap();
    let mut file = File::open(&file_name).unwrap();
    let mut buf = vec![];
    file.read_to_end(&mut buf).unwrap();

    let mut code = String::from("");
    let res = read_docx(&buf).unwrap().document;
    for document_child in res.children {
        match document_child {
            DocumentChild::Paragraph(paragraph) => {
                for paragraph_child in paragraph.children {
                    match paragraph_child {
                        ParagraphChild::Run(run) => {
                            for run_child in run.children {
                                match run_child {
                                    RunChild::Tab(_) => {
                                        code.push_str("    ");
                                    }
                                    RunChild::Text(text) => {
                                        code.push_str(&text.text);
                                        code.push('\n');
                                    }
                                    _ => {}
                                }
                            }
                        }
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }

    code = code.replace("\u{201c}", "\"");
    code = code.replace("\u{201d}", "\"");
    code = code.replace("\u{2018}", "\'");
    code = code.replace("\u{2019}", "\'");

    pyo3::prepare_freethreaded_python();
    let py = pyo3::Python::acquire_gil();
    py.python().run(&code, None, None).unwrap();
}
