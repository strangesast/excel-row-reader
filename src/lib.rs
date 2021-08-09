#![deny(clippy::all)]

#[macro_use]
extern crate napi_derive;

use calamine::vba::VbaProject;
use calamine::{DataType, Error, Metadata, Range, Reader, Xls, Xlsb, Xlsx};
use napi::Error as NapiError;
use napi::Result as NapiResult;
use napi::{
  CallContext, JsBuffer, JsBufferValue, JsNumber, JsObject, JsString, JsUnknown, Status, ValueType,
};
use std::borrow::Cow;
use std::collections::HashMap;
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
  let err = Err(NapiError::new(
    Status::InvalidArg,
    format!("not in {} format", s),
  ));
  return match s {
    "xlsx" | "xlsm" => Xlsx::new(cur).map_or(err, |wb| Ok(CurSheets::Xlsx(wb))),
    "xlsb" => Xlsb::new(cur).map_or(err, |wb| Ok(CurSheets::Xlsb(wb))),
    "xls" => Xls::new(cur).map_or(err, |wb| Ok(CurSheets::Xls(wb))),
    _ => Err(NapiError::new(Status::InvalidArg, String::new())),
  };
}

#[js_function(4)]
fn parse(ctx: CallContext) -> NapiResult<JsObject> {
  let u = ctx.get::<JsString>(0)?.into_utf8()?;
  let s = u.as_str()?;
  let buf = &mut ctx.get::<JsBuffer>(1)?.into_value()?;
  let headers = ctx.get::<JsObject>(2)?;
  let sheet_input = ctx.get::<JsUnknown>(3)?;

  let mut wb = get_wb(s, buf)?;

  // if sheet is string, get sheet by name, if number, get sheet by index, if
  // undefined, get first sheet
  let range_result = match sheet_input.get_type()? {
    ValueType::String => {
      let sheet_name_o: JsString = unsafe { sheet_input.cast() };
      let sheet_name = sheet_name_o.into_utf8()?;
      let sheet_name_str = sheet_name.as_str()?;
      let r = wb.worksheet_range(&sheet_name_str).unwrap().unwrap();
      Ok(r)
    }
    ValueType::Number => {
      let sheet_number_o: JsNumber = unsafe { sheet_input.cast() };
      let sheet_number = sheet_number_o.get_int32()? as usize;
      let r = wb.worksheet_range_at(sheet_number).unwrap().unwrap();
      Ok(r)
    }
    ValueType::Null | ValueType::Undefined => {
      let r = wb.worksheet_range_at(0).unwrap().unwrap();
      Ok(r)
    }
    _ => Err(NapiError::new(Status::InvalidArg, String::new())),
  };

  let range = range_result?;
  let header_vec = read_headers(headers)?;

  return parse_range(ctx, range, header_vec);
}

// convert [string, string][] to Vec<(trimmed,uppercased,String,String)>
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
        .to_uppercase()
        .trim()
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

fn parse_range(
  ctx: CallContext,
  range: Range<DataType>,
  headers: Vec<(String, String)>,
) -> NapiResult<JsObject> {
  let (h, w) = range.get_size();
  assert!(
    w > 1 && h > 1,
    "invalid range: must have >1 row and >0 columns"
  );
  let mut output = ctx.env.create_array_with_length(h - 1)?; // array of arrays to return to node

  let header_map: HashMap<String, usize> = (0..w)
    .filter_map(|i| match range.get((0, i)) {
      Some(cell) => Some((cell.to_string().to_uppercase().trim().to_string(), i)),
      None => None,
    })
    .collect();

  let indexes: Vec<(usize, JsString)> = headers
    .iter()
    .map(|(s0, s1)| match header_map.get(s0) {
      Some(&index) => (index, ctx.env.create_string(s1).unwrap()),
      None => panic!("missing header in file"),
    })
    .collect();
  // indexes.sort_by_key(|k| k.0);

  for j in 1..h {
    let mut out = ctx.env.create_object()?;

    for (i, key) in indexes.iter() {
      // for k in 0..l {
      //   let (i, key) = indexes[k];
      match range.get((j, *i)) {
        Some(c) => match *c {
          DataType::String(ref s) => {
            let ss = s.trim();
            if ss != "" {
              out.set_property(*key, ctx.env.create_string(ss)?)?;
            }
          }
          DataType::DateTime(ref f) | DataType::Float(ref f) => {
            out.set_property(*key, ctx.env.create_double(*f)?)?;
          }
          DataType::Int(ref d) => {
            out.set_property(*key, ctx.env.create_int64(*d)?)?;
          }
          DataType::Bool(ref b) => {
            out.set_property(*key, ctx.env.get_boolean(*b)?)?;
          }
          _ => {}
        },
        _ => {}
      }
    }
    output.set_element((j - 1) as u32, out)?;
  }
  return Ok(output);
}
