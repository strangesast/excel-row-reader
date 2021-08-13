#![deny(clippy::all)]

#[macro_use]
extern crate napi_derive;

use calamine::vba::VbaProject;
use calamine::{DataType, Error, Metadata, Range, Reader, Xls, Xlsb, Xlsx};
use napi::Error as NapiError;
use napi::Result as NapiResult;
use napi::{CallContext, JsBuffer, JsBufferValue, JsObject, JsString, Status};
use std::borrow::Cow;
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
fn init(mut exports: JsObject) -> NapiResult<()> {
  exports.create_named_method("parse", parse)?;
  Ok(())
}

// pattern copied from calamine::Sheets
enum CurSheets {
  Xls(Xls<Cursor<Vec<u8>>>),
  Xlsx(Xlsx<Cursor<Vec<u8>>>),
  Xlsb(Xlsb<Cursor<Vec<u8>>>),
}

impl Reader for CurSheets {
  type RS = Cursor<Vec<u8>>;
  type Error = Error;

  /// Creates a new instance.
  fn new(_reader: Self::RS) -> Result<Self, Self::Error> {
    Err(Error::Msg("Sheets must be created from a Path"))
  }

  /// Gets `VbaProject`
  fn vba_project(&mut self) -> Option<Result<Cow<'_, VbaProject>, Self::Error>> {
    Some(Err(Error::Msg("not implemented")))
  }

  /// Initialize
  fn metadata(&self) -> &Metadata {
    match *self {
      CurSheets::Xls(ref e) => e.metadata(),
      CurSheets::Xlsx(ref e) => e.metadata(),
      CurSheets::Xlsb(ref e) => e.metadata(),
    }
  }

  /// Read worksheet data in corresponding worksheet path
  fn worksheet_range(&mut self, name: &str) -> Option<Result<Range<DataType>, Self::Error>> {
    match *self {
      CurSheets::Xls(ref mut e) => e.worksheet_range(name).map(|r| r.map_err(Error::Xls)),
      CurSheets::Xlsx(ref mut e) => e.worksheet_range(name).map(|r| r.map_err(Error::Xlsx)),
      CurSheets::Xlsb(ref mut e) => e.worksheet_range(name).map(|r| r.map_err(Error::Xlsb)),
    }
  }

  /// Read worksheet formula in corresponding worksheet path
  fn worksheet_formula(&mut self, _name: &str) -> Option<Result<Range<String>, Self::Error>> {
    Some(Err(Error::Msg("not implemented")))
  }

  fn worksheets(&mut self) -> Vec<(String, Range<DataType>)> {
    match *self {
      CurSheets::Xls(ref mut e) => e.worksheets(),
      CurSheets::Xlsx(ref mut e) => e.worksheets(),
      CurSheets::Xlsb(ref mut e) => e.worksheets(),
    }
  }
}

fn get_wb(s: &str, buf: &mut JsBufferValue) -> NapiResult<CurSheets> {
  let cur = Cursor::new(buf.to_vec());
  return match s {
    "xlsx" => Ok(CurSheets::Xlsx(Xlsx::new(cur).unwrap())),
    "xlsb" => Ok(CurSheets::Xlsb(Xlsb::new(cur).unwrap())),
    "xls" => Ok(CurSheets::Xls(Xls::new(cur).unwrap())),
    _ => Err(NapiError::new(Status::InvalidArg, String::new())),
  };
}

#[js_function(4)]
fn parse(ctx: CallContext) -> NapiResult<JsObject> {
  let u = ctx.get::<JsString>(0)?.into_utf8()?;
  let s = u.as_str()?;
  let buf = &mut ctx.get::<JsBuffer>(1)?.into_value()?;
  let sheet_name = ctx.get::<JsString>(2)?.into_utf8()?;
  let headers = ctx.get::<JsObject>(3)?;

  let header_vec = read_headers(headers)?;

  // let cur = Cursor::new(buf.to_vec());
  // let mut wb = Xlsb::new(cur).unwrap();
  let mut wb = get_wb(s, buf)?;

  // let sheets = wb.sheet_names();
  // let mut obj = ctx.env.create_array_with_length(sheets.len())?;
  // for (i, sheet) in sheets.iter().enumerate() {
  //   obj.set_element(i as u32, ctx.env.create_string(sheet)?)?;
  // }

  let sheet_name_str = sheet_name.as_str()?;
  let r = wb.worksheet_range(&sheet_name_str).unwrap();
  assert!(r.is_ok(), "missing sheet '{}'", sheet_name_str);
  let range = r.unwrap();
  return parse_range(ctx, range, header_vec);
}

fn read_headers(inp: JsObject) -> NapiResult<Vec<(String, String)>> {
  assert!(inp.is_array()?);

  let m = inp.get_array_length()?;
  let mut header_vec = Vec::new();
  for i in 0..m {
    let el = inp.get_element::<JsObject>(i)?;
    assert!(el.is_array()?, "invalid element {} must be array", i);

    let s = (
      el.get_element::<JsString>(0)?
        .into_utf8()?
        .as_str()?
        .to_string(),
      el.get_element::<JsString>(1)?
        .into_utf8()?
        .as_str()?
        .to_string(),
    );
    header_vec.push(s);
  }
  return Ok(header_vec);
}

/**
 * headers is case-insensitive list of strings to match against first row
 *
 */
fn parse_range(
  ctx: CallContext,
  range: Range<DataType>,
  output_headers: Vec<(String, String)>,
) -> NapiResult<JsObject> {
  let l = range.get_size().0;
  let m = output_headers.len();

  let mut output = ctx.env.create_array_with_length(l - 1)?; // array of arrays to return to node

  let mut it = range.rows();

  let header_values: Vec<String> = it
    .nth(0)
    .unwrap()
    .iter()
    .map(|c| c.to_string().to_uppercase().trim().to_string())
    .collect();

  let indexes: Vec<(usize, JsString)> = output_headers
    .iter()
    .map(|s| {
      (
        header_values
          .iter()
          .position(|ss| s.0.to_uppercase().eq(ss))
          .unwrap(),
        ctx.env.create_string_from_std(s.1.clone()).unwrap(),
      )
    })
    .collect();

  let mut ind = indexes.clone();
  ind.sort_by_key(|k| k.0);

  // let positions: Vec<usize> = ind
  //   .iter()
  //   .map(|i| indexes.iter().position(|j| j == i).unwrap())
  //   .collect();

  // let header_strings: Vec<JsString> = header_values
  //   .iter()
  //   .map(|s| ctx.env.create_string_from_std(s.clone()).unwrap())
  //   .collect();

  for (j, row) in it.enumerate() {
    // let mut out = ctx.env.create_array_with_length(m)?;
    let mut out = ctx.env.create_object()?;
    let mut r = row.iter().enumerate();

    let mut i = 0usize;
    while i < m {
      if let Some(t) = r.next() {
        let (jj, c) = t;
        if jj == ind[i].0 {
          // let k = positions[i] as u32;
          // let k = header_strings[positions[i]];
          let k = ind[i].1;
          match *c {
            DataType::String(ref s) => {
              // out.set_element(k, ctx.env.create_string(s.as_str())?)?;
              if s.trim() != "" {
                out.set_property(k, ctx.env.create_string(s.as_str())?)?;
              }
            }
            DataType::Float(ref f) => {
              // includes dates
              // out.set_element(k, ctx.env.create_double(*f)?)?;
              out.set_property(k, ctx.env.create_double(*f)?)?;
            }
            DataType::Int(ref d) => {
              // out.set_element(k, ctx.env.create_int64(*d)?)?;
              out.set_property(k, ctx.env.create_int64(*d)?)?;
            }
            DataType::Bool(ref b) => {
              // out.set_element(k, ctx.env.get_boolean(*b)?)?;
              out.set_property(k, ctx.env.get_boolean(*b)?)?;
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
