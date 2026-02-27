use std::mem::ManuallyDrop;

use xll_rs::convert::build_multi;
use xll_rs::memory::xlAutoFree12;
use xll_rs::types::*;

const WARMUP_ITERATIONS: usize = 100;
const STRESS_ITERATIONS: usize = 5000;
const LOG_INTERVAL: usize = 500;
const MAX_HEAP_DELTA_KB: i64 = 1024;

// ============================================================================
// Process Memory Tracking (Windows)
// ============================================================================

#[repr(C)]
#[allow(non_snake_case)]
struct PROCESS_MEMORY_COUNTERS {
    cb: u32,
    PageFaultCount: u32,
    PeakWorkingSetSize: usize,
    WorkingSetSize: usize,
    QuotaPeakPagedPoolUsage: usize,
    QuotaPagedPoolUsage: usize,
    QuotaPeakNonPagedPoolUsage: usize,
    QuotaNonPagedPoolUsage: usize,
    PagefileUsage: usize,
    PeakPagefileUsage: usize,
}

extern "system" {
    fn GetCurrentProcess() -> *mut std::ffi::c_void;
    fn K32GetProcessMemoryInfo(
        process: *mut std::ffi::c_void,
        ppsmemCounters: *mut PROCESS_MEMORY_COUNTERS,
        cb: u32,
    ) -> i32;
}

#[derive(Clone, Copy)]
struct MemSnapshot {
    pagefile: usize,
    working_set: usize,
    peak_working_set: usize,
}

impl MemSnapshot {
    fn now() -> Self {
        let mut counters = PROCESS_MEMORY_COUNTERS {
            cb: std::mem::size_of::<PROCESS_MEMORY_COUNTERS>() as u32,
            PageFaultCount: 0,
            PeakWorkingSetSize: 0,
            WorkingSetSize: 0,
            QuotaPeakPagedPoolUsage: 0,
            QuotaPagedPoolUsage: 0,
            QuotaPeakNonPagedPoolUsage: 0,
            QuotaNonPagedPoolUsage: 0,
            PagefileUsage: 0,
            PeakPagefileUsage: 0,
        };
        unsafe {
            K32GetProcessMemoryInfo(
                GetCurrentProcess(),
                &mut counters,
                counters.cb,
            );
        }
        Self {
            pagefile: counters.PagefileUsage,
            working_set: counters.WorkingSetSize,
            peak_working_set: counters.PeakWorkingSetSize,
        }
    }

    fn display(&self) -> String {
        format!(
            "heap: {} KB, working_set: {} KB, peak: {} KB",
            self.pagefile / 1024,
            self.working_set / 1024,
            self.peak_working_set / 1024,
        )
    }

    fn heap_delta_from(&self, baseline: &MemSnapshot) -> i64 {
        self.pagefile as i64 - baseline.pagefile as i64
    }
}

fn log_mem(label: &str, iter: usize, baseline: &MemSnapshot) {
    let now = MemSnapshot::now();
    let delta = now.heap_delta_from(baseline);
    let sign = if delta >= 0 { "+" } else { "" };
    eprintln!(
        "  [{}] iter {:>4} | {} | delta: {}{} KB",
        label,
        iter,
        now.display(),
        sign,
        delta / 1024,
    );
}

fn boxed_dll_free(mut oper: XLOPER12) -> *mut XLOPER12 {
    oper.xltype |= XLBIT_DLL_FREE;
    Box::into_raw(Box::new(oper))
}

fn stress_type<F>(label: &str, make: F)
where
    F: Fn() -> *mut XLOPER12,
{
    for _ in 0..WARMUP_ITERATIONS {
        let p = make();
        xlAutoFree12(p);
    }

    let baseline = MemSnapshot::now();
    eprintln!();
    log_mem(label, 0, &baseline);

    for i in 1..=STRESS_ITERATIONS {
        let p = make();
        xlAutoFree12(p);

        if i % LOG_INTERVAL == 0 {
            log_mem(label, i, &baseline);
        }
    }

    let final_snap = MemSnapshot::now();
    let delta_kb = final_snap.heap_delta_from(&baseline) / 1024;
    log_mem(label, STRESS_ITERATIONS, &baseline);
    eprintln!(
        "  [{}] final heap delta: {} KB over {} iterations",
        label, delta_kb, STRESS_ITERATIONS
    );

    assert!(
        delta_kb < MAX_HEAP_DELTA_KB,
        "Heap grew by {} KB over {} iterations — possible memory leak",
        delta_kb,
        STRESS_ITERATIONS
    );
}

fn make_multi_numeric() -> *mut XLOPER12 {
    let cells = vec![
        XLOPER12::from_f64(1.0),
        XLOPER12::from_f64(2.0),
        XLOPER12::from_f64(3.0),
        XLOPER12::from_f64(4.0),
    ];
    build_multi(cells, 2, 2)
}

fn make_multi_mixed() -> *mut XLOPER12 {
    let cells = vec![
        XLOPER12::from_str("A"),
        XLOPER12::from_f64(1.0),
        XLOPER12::from_str("B"),
        XLOPER12::from_err(XLERR_NA),
    ];
    build_multi(cells, 2, 2)
}

fn make_sref() -> *mut XLOPER12 {
    boxed_dll_free(XLOPER12 {
        val: XLOPER12Val {
            sref: ManuallyDrop::new(XLOPER12SRef {
                count: 1,
                ref_: XLREF12 {
                    rw_first: 0,
                    rw_last: 0,
                    col_first: 0,
                    col_last: 0,
                },
            }),
        },
        xltype: XLTYPE_SREF,
    })
}

fn make_ref() -> *mut XLOPER12 {
    boxed_dll_free(XLOPER12 {
        val: XLOPER12Val {
            mref: ManuallyDrop::new(XLOPER12MRef {
                lpmref: std::ptr::null_mut(),
                id_sheet: 0,
            }),
        },
        xltype: XLTYPE_REF,
    })
}

fn make_flow() -> *mut XLOPER12 {
    boxed_dll_free(XLOPER12 {
        val: XLOPER12Val {
            flow: ManuallyDrop::new(XLOPER12Flow {
                valflow: XLOPER12FlowVal { level: 1 },
                rw: 0,
                col: 0,
                xlflow: XLFLOW_RESTART,
            }),
        },
        xltype: XLTYPE_FLOW,
    })
}

fn make_bigdata() -> *mut XLOPER12 {
    boxed_dll_free(XLOPER12 {
        val: XLOPER12Val {
            bigdata: ManuallyDrop::new(XLOPER12BigData {
                h: XLOPER12BigDataHandle {
                    lpb_data: std::ptr::null_mut(),
                },
                cb_data: 0,
            }),
        },
        xltype: XLTYPE_BIGDATA,
    })
}

#[test]
fn leak_num() {
    stress_type("num", || boxed_dll_free(XLOPER12::from_f64(3.14)));
}

#[test]
fn leak_int() {
    stress_type("int", || boxed_dll_free(XLOPER12::from_int(7)));
}

#[test]
fn leak_bool() {
    stress_type("bool", || boxed_dll_free(XLOPER12::from_bool(true)));
}

#[test]
fn leak_err() {
    stress_type("err", || boxed_dll_free(XLOPER12::from_err(XLERR_NUM)));
}

#[test]
fn leak_nil() {
    stress_type("nil", || boxed_dll_free(XLOPER12::nil()));
}

#[test]
fn leak_missing() {
    stress_type("missing", || boxed_dll_free(XLOPER12::missing()));
}

#[test]
fn leak_str() {
    stress_type("str", || boxed_dll_free(XLOPER12::from_str("hello")));
}

#[test]
fn leak_blank() {
    stress_type("blank", || boxed_dll_free(XLOPER12::from_str("")));
}

#[test]
fn leak_multi_numeric() {
    stress_type("multi_num", make_multi_numeric);
}

#[test]
fn leak_multi_mixed() {
    stress_type("multi_mixed", make_multi_mixed);
}

#[test]
fn leak_sref() {
    stress_type("sref", make_sref);
}

#[test]
fn leak_ref() {
    stress_type("ref", make_ref);
}

#[test]
fn leak_flow() {
    stress_type("flow", make_flow);
}

#[test]
fn leak_bigdata() {
    stress_type("bigdata", make_bigdata);
}
