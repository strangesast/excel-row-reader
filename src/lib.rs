#![deny(clippy::all)]

#[macro_use]
extern crate napi_derive;

use calamine::{DataType, Range, Reader, Xlsb};
use napi::{CallContext, JsBuffer, JsObject, JsString, Result};
use std::io::Cursor;

#[cfg(all(
  any(windows, unix),
  target_arch = "x86_64",
  not(target_env = "musl"),
  not(debug_assertions)
))]
#[global_allocator]
static ALLOC: mimalloc::MiMalloc = mimalloc::MiMalloc;

#[module_exports]
fn init(mut exports: JsObject) -> Result<()> {
  exports.create_named_method("world", world_xlsb)?;
  Ok(())
}

#[js_function(3)]
fn world_xlsb(ctx: CallContext) -> Result<JsObject> {
  let buf = &mut ctx.get::<JsBuffer>(0)?.into_value()?;
  let sheet_name = ctx.get::<JsString>(1)?.into_utf8()?;
  let headers = ctx.get::<JsObject>(2)?;

  let header_vec = read_headers(headers)?;

  let mut wb = Xlsb::new(Cursor::new(buf.to_vec())).unwrap();

  // let sheets = wb.sheet_names();
  // let mut obj = ctx.env.create_array_with_length(sheets.len())?;
  // for (i, sheet) in sheets.iter().enumerate() {
  //   obj.set_element(i as u32, ctx.env.create_string(sheet)?)?;
  // }

  let sheet_name_str = sheet_name.as_str()?;
  let r = wb.worksheet_range(&sheet_name_str).unwrap();
  assert!(r.is_ok(), "missing sheet '{}'", sheet_name_str);
  let range = r.unwrap();
  return world_wb(ctx, range, header_vec);
}

fn read_headers(inp: JsObject) -> Result<Vec<String>> {
  assert!(inp.is_array()?);

  let m = inp.get_array_length()? as usize;
  let mut header_vec = Vec::new();
  for i in 0..(m as u32) {
    let el: JsString = inp.get_element(i)?;
    let s = el.into_utf8()?;
    header_vec.push(String::from(s.as_str()?));
  }
  return Ok(header_vec);
}

/**
 * headers is case-insensitive list of strings to match against first row
 *
 */
fn world_wb(ctx: CallContext, range: Range<DataType>, headers: Vec<String>) -> Result<JsObject> {
  let l = range.get_size().0;
  let m = headers.len();

  let mut output = ctx.env.create_array_with_length(l - 1)?; // array of arrays to return to node

  let mut it = range.rows();
  let header = it.nth(0).unwrap(); // TODO handle error

  let header_values: Vec<String> = header
    .iter()
    .map(|c| c.to_string().to_uppercase().trim().to_string())
    .collect();

  let indexes: Vec<usize> = headers
    .iter()
    .map(|s| {
      header_values
        .iter()
        .position(|ss| s.to_uppercase().eq(ss))
        .unwrap()
    })
    .collect();

  let mut ind: Vec<usize> = indexes.clone();
  ind.sort();

  let positions: Vec<usize> = ind
    .iter()
    .map(|i| indexes.iter().position(|j| j == i).unwrap())
    .collect();

  for (j, row) in it.enumerate() {
    let mut out = ctx.env.create_array_with_length(m)?;
    let mut r = row.iter().enumerate();

    let mut i = 0usize;
    while i < m {
      if let Some(t) = r.next() {
        let (jj, c) = t;
        if jj == ind[i] {
          let k = positions[i] as u32;
          match *c {
            DataType::String(ref s) => {
              out.set_element(k, ctx.env.create_string(s.as_str())?)?;
            }
            DataType::DateTime(ref f) | DataType::Float(ref f) => {
              // let d = c.as_date().unwrap().format("%Y-%m-%d").to_string();
              // out.set_element(ii, ctx.env.create_string(d.as_str())?)?;
              out.set_element(k, ctx.env.create_double(*f)?)?;
            }
            DataType::Int(ref d) => {
              out.set_element(k, ctx.env.create_int64(*d)?)?;
            }
            DataType::Bool(ref b) => {
              out.set_element(k, ctx.env.get_boolean(*b)?)?;
            }
            _ => {}
          }
          i += 1;
          continue;
        }
      } else {
        break;
      }
    }
    output.set_element(j as u32, out)?;
  }
  return Ok(output);
}
