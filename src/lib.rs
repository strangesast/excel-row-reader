#![deny(clippy::all)]

#[macro_use]
extern crate napi_derive;

use calamine::vba::VbaProject;
use calamine::Error as CalamineError;
use calamine::{DataType, Metadata, Range, Reader, Xls, Xlsb, Xlsx};
use napi::Error as NapiError;
use napi::Result as NapiResult;
use napi::{CallContext, JsBuffer, JsNumber, JsObject, JsString, JsUnknown, Status, ValueType};
use std::borrow::Cow;
use std::collections::HashMap;
use std::io::Cursor;

use json::JsonValue;
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
  exports.create_named_method("dump", dump)?;
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
  type Error = CalamineError;

  /// Creates a new instance.
  fn new(_reader: Self::RS) -> Result<Self, Self::Error> {
    Err(CalamineError::Msg("Sheets must be created from a Path"))
  }

  /// Gets `VbaProject`
  fn vba_project(&mut self) -> Option<Result<Cow<'_, VbaProject>, Self::Error>> {
    Some(Err(CalamineError::Msg("not implemented")))
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
      CurSheets::Xls(ref mut e) => e
        .worksheet_range(name)
        .map(|r| r.map_err(CalamineError::Xls)),
      CurSheets::Xlsx(ref mut e) => e
        .worksheet_range(name)
        .map(|r| r.map_err(CalamineError::Xlsx)),
      CurSheets::Xlsb(ref mut e) => e
        .worksheet_range(name)
        .map(|r| r.map_err(CalamineError::Xlsb)),
    }
  }

  /// Read worksheet formula in corresponding worksheet path
  fn worksheet_formula(&mut self, _name: &str) -> Option<Result<Range<String>, Self::Error>> {
    Some(Err(CalamineError::Msg("not implemented")))
  }

  fn worksheets(&mut self) -> Vec<(String, Range<DataType>)> {
    match *self {
      CurSheets::Xls(ref mut e) => e.worksheets(),
      CurSheets::Xlsx(ref mut e) => e.worksheets(),
      CurSheets::Xlsb(ref mut e) => e.worksheets(),
    }
  }
}

// if sheet is string, get sheet by name, if number, get sheet by index, if
// undefined, get first sheet
fn get_range_from_sheet(
  wb: &mut CurSheets,
  sheet_input: JsUnknown,
) -> Result<Range<DataType>, &'static str> {
  let r: Option<Range<DataType>> = match sheet_input.get_type().unwrap_or(ValueType::Undefined) {
    ValueType::String => {
      let sheet_name_o: JsString = unsafe { sheet_input.cast() };
      let sheet_name = sheet_name_o.into_utf8().unwrap();
      let sheet_name_str = sheet_name.as_str().unwrap();
      let r = wb.worksheet_range(&sheet_name_str);
      r.map_or(None, |rr| Some(rr.unwrap()))
    }
    ValueType::Number => {
      let sheet_number_o: JsNumber = unsafe { sheet_input.cast() };
      let sheet_number = sheet_number_o.get_int32().unwrap() as usize;
      let r = wb.worksheet_range_at(sheet_number);
      r.map_or(None, |rr| Some(rr.unwrap()))
    }
    ValueType::Null | ValueType::Undefined => {
      let r = wb.worksheet_range_at(0);
      r.map_or(None, |rr| Some(rr.unwrap()))
    }
    _ => None,
  };

  match r {
    Some(rr) => Ok(rr),
    None => Err("sheet not found"),
  }
}

#[js_function(4)]
fn dump(ctx: CallContext) -> NapiResult<JsObject> {
  let u = ctx.get::<JsString>(0)?.into_utf8()?;
  let s = u.as_str()?;
  let buf = &mut ctx.get::<JsBuffer>(1)?.into_value()?;
  let headers = ctx.get::<JsObject>(2)?;
  let sheet_input = ctx.get::<JsUnknown>(3)?;

  let cur = Cursor::new(buf.to_vec());
  let err = Err(NapiError::new(
    Status::InvalidArg,
    format!("not in {} format", s),
  ));
  let mut wb = (match s {
    "xlsx" | "xlsm" => Xlsx::new(cur).map_or(err, |wb| Ok(CurSheets::Xlsx(wb))),
    "xlsb" => Xlsb::new(cur).map_or(err, |wb| Ok(CurSheets::Xlsb(wb))),
    "xls" => Xls::new(cur).map_or(err, |wb| Ok(CurSheets::Xls(wb))),
    _ => Err(NapiError::new(
      Status::InvalidArg,
      String::from("must be \"xlsx\", \"xlsm\", \"xlsb\", or \"xls\""),
    )),
  })?;

  let range = get_range_from_sheet(&mut wb, sheet_input).map_err(|_| {
    NapiError::new(
      Status::InvalidArg,
      String::from("failed to get range from sheet"),
    )
  })?;
  let header_vec = read_headers(headers)?;

  return dump_range(ctx, range, header_vec);
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
      el.get_element::<JsUnknown>(0)?
        .coerce_to_string()?
        .into_utf8()?
        .as_str()?
        .to_uppercase()
        .trim()
        .to_string(),
      el.get_element::<JsUnknown>(1)?
        .coerce_to_string()?
        .into_utf8()?
        .as_str()?
        .to_string(),
    );
    header_vec.push(s);
  }
  return Ok(header_vec);
}

fn dump_range(
  ctx: CallContext,
  range: Range<DataType>,
  headers: Vec<(String, String)>,
) -> NapiResult<JsObject> {
  // opens a new workbook
  let (h, w) = range.get_size();
  let mut output = JsonValue::new_array();

  let header_map: HashMap<String, usize> = (0..w)
    .filter_map(|i| match range.get((0, i)) {
      Some(cell) => Some((cell.to_string().to_uppercase().trim().to_string(), i)),
      None => None,
    })
    .collect();

  let mut indexes: Vec<(usize, String)> = headers
    .iter()
    .map(
      |(s0, s1)| match header_map.get(&s0.to_uppercase().trim().to_string()) {
        Some(&index) => (index, s1.to_string()),
        None => panic!("missing header in file \"{}\"", s0),
      },
    )
    .collect();
  indexes.sort_by_key(|k| k.0);

  for j in 1..h {
    let mut out = JsonValue::new_object();

    for (i, key) in indexes.iter() {
      match range.get((j, *i)) {
        Some(c) => match *c {
          DataType::String(ref s) => {
            let ss = s.trim();
            if ss != "" {
              out.insert(key, ss).unwrap();
            }
          }
          DataType::DateTime(ref f) | DataType::Float(ref f) => {
            out.insert(key, JsonValue::from(*f)).unwrap();
          }
          DataType::Int(ref d) => {
            out.insert(key, JsonValue::from(*d)).unwrap();
          }
          DataType::Bool(ref b) => {
            out.insert(key, JsonValue::from(*b)).unwrap();
          }
          _ => {}
        },
        _ => {}
      }
    }
    output.push(out).unwrap()
  }

  let out = ctx
    .env
    .create_buffer_with_data(output.dump().into_bytes())?;

  let o = out.into_raw().coerce_to_object()?;

  return Ok(o);
}
