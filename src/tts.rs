use std::ops::Drop;

use anyhow::ensure;
use winapi::{
    shared::{guiddef::GUID, winerror::SUCCEEDED},
    um::{
        combaseapi::{
            CLSIDFromProgID, CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_ALL,
        },
        objbase::COINIT_MULTITHREADED,
        sapi51::{ISpVoice, SPF_ASYNC, SPF_DEFAULT},
    },
    Interface,
};
use wio::{com::ComPtr, wide::ToWide};

pub struct Speaker {
    _com_init: ComInitRaii,
    voice: ComPtr<ISpVoice>,
    buffer: Vec<u16>,
}

impl Speaker {
    pub fn new() -> anyhow::Result<Speaker> {
        let com_init = ComInitRaii::new()?;
        let clsid = get_sapi_voice_clsid()?;

        let mut ptr = std::ptr::null_mut::<ISpVoice>();
        let hr = unsafe {
            CoCreateInstance(
                &clsid,
                std::ptr::null_mut(),
                CLSCTX_ALL,
                &ISpVoice::uuidof(),
                (&mut ptr) as *mut _ as _,
            )
        };
        ensure!(SUCCEEDED(hr), "Failed to initialize SAPI Voice instance");
        let voice = unsafe { ComPtr::from_raw(ptr) };

        Ok(Speaker {
            _com_init: com_init,
            voice,
            buffer: Vec::with_capacity(1024),
        })
    }

    pub fn speak(&mut self, text: &str) -> anyhow::Result<()> {
        self.do_speak(text, SPF_DEFAULT)
    }

    pub fn speak_async(&mut self, text: &str) -> anyhow::Result<()> {
        self.do_speak(text, SPF_ASYNC)
    }

    fn do_speak(&mut self, text: &str, flags: u32) -> anyhow::Result<()> {
        self.buffer.clear();
        self.buffer.extend(text.encode_utf16());
        self.buffer.push(0);

        let hr = unsafe { self.voice.Speak(self.buffer.as_ptr(), flags, 0 as _) };
        ensure!(SUCCEEDED(hr), "Failed to speak");

        Ok(())
    }
}

unsafe impl Send for Speaker {}

struct ComInitRaii {
    _priv: (),
}

impl ComInitRaii {
    pub fn new() -> anyhow::Result<Self> {
        let hr = unsafe { CoInitializeEx(0 as _, COINIT_MULTITHREADED) };
        ensure!(SUCCEEDED(hr), "CoInitializeEx failed");

        Ok(ComInitRaii { _priv: () })
    }
}

impl Drop for ComInitRaii {
    fn drop(&mut self) {
        unsafe {
            CoUninitialize();
        }
    }
}

fn get_sapi_voice_clsid() -> anyhow::Result<GUID> {
    let progid = "SAPI.SpVoice".to_wide_null();
    let mut clsid = GUID::default();
    let hr = unsafe { CLSIDFromProgID(progid.as_ptr(), &mut clsid) };
    ensure!(SUCCEEDED(hr), "Failed to find SAPI.SpVoice class");

    Ok(clsid)
}
