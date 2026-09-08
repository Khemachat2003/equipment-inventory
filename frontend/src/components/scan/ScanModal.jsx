import { useEffect, useRef, useState } from 'react';
import { BrowserMultiFormatReader } from '@zxing/browser';
import { BarcodeFormat } from '@zxing/library';
import Icon from '../ui/Icon.jsx';

const FORMATS = [
  BarcodeFormat.QR_CODE,
  BarcodeFormat.CODE_128,
  BarcodeFormat.CODE_39,
  BarcodeFormat.DATA_MATRIX,
  BarcodeFormat.EAN_13,
  BarcodeFormat.EAN_8,
  BarcodeFormat.CODABAR,
];

export default function ScanModal({ onClose, onResult }) {
  const videoRef = useRef(null);
  const controlsRef = useRef(null);
  const [error, setError] = useState('');
  const [manual, setManual] = useState('');
  const [starting, setStarting] = useState(true);

  function stop() {
    try { controlsRef.current?.stop?.(); } catch (e) {}
    try { BrowserMultiFormatReader.releaseAllStreams(); } catch (e) {}
    controlsRef.current = null;
  }

  function handleFound(text) {
    const t = (text || '').trim();
    if (!t || !onResult) return;
    stop();
    onResult(t);
  }

  async function start() {
    setStarting(true);
    setError('');
    try {
      const reader = new BrowserMultiFormatReader();
      reader.possibleFormats = FORMATS;
      const controls = await reader.decodeFromVideoDevice(undefined, videoRef.current, (result) => {
        if (result && result.getText) handleFound(result.getText());
      });
      controlsRef.current = controls;
    } catch (e) {
      console.error('scan start error:', e);
      setError('เปิดกล้องไม่ได้ — ตรวจสิทธิ์กล้อง (หรือใช้ช่อง "พิมพ์ Serial" ด้านล่างแทน)');
    } finally {
      setStarting(false);
    }
  }

  useEffect(() => {
    start();
    return stop;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  function submitManual() {
    const t = manual.trim();
    if (!t) return;
    stop();
    onResult && onResult(t);
  }

  return (
    <div
      className="fixed inset-0 z-50 flex items-start justify-center p-4 bg-black/40 backdrop-blur-sm overflow-y-auto"
      onClick={onClose}
    >
      <div
        className="w-full max-w-md bg-white rounded-2xl shadow-xl my-6 overflow-hidden"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between px-5 py-4 border-b border-[var(--g100)]">
          <div className="flex items-center gap-2">
            <span className="text-[var(--blue)]"><Icon name="document_scanner" /></span>
            <span className="text-[15px] font-bold text-[var(--text)]">สแกน Barcode / Serial</span>
          </div>
          <button
            onClick={onClose}
            className="w-8 h-8 flex items-center justify-center rounded-lg text-[var(--tmuted)] hover:bg-[var(--surface2)]"
          >
            <Icon name="close" size="sm" />
          </button>
        </div>

        <div className="p-5 space-y-3">
          {error ? (
            <div className="flex items-start gap-2.5 px-3 py-3 rounded-lg bg-[var(--red-l)] border border-[var(--red-b)] text-[13px] text-[var(--red)]">
              <Icon name="videocam_off" size="sm" />
              <div>
                <div className="font-semibold">ไม่สามารถเปิดกล้องได้</div>
                <div className="text-[12px] text-[var(--tsub)]">ใช้ช่องพิมพ์ด้านล่างแทน หรือลองอีกครั้ง</div>
              </div>
            </div>
          ) : (
            <div className="relative rounded-xl overflow-hidden bg-black">
              <video ref={videoRef} className="w-full h-64 object-cover" muted playsInline />
              <div className="absolute inset-0 pointer-events-none flex items-center justify-center">
                <div className="w-4/5 h-28 border-2 border-white/70 rounded-xl" />
              </div>
              {starting && (
                <div className="absolute inset-0 flex flex-col items-center justify-center gap-2 text-white/90 text-[13px] bg-black/50">
                  <Icon name="hourglass_top" className="animate-spin" />
                  กำลังเปิดกล้อง...
                </div>
              )}
            </div>
          )}

          <div className="text-center text-[12px] text-[var(--tmuted)]">
            วาง Barcode / QR Code ให้อยู่ในกรอบด้านซ้ายล่างของตัวเครื่อง
          </div>

          <div className="flex items-center gap-2 py-1">
            <div className="flex-1 h-px bg-[var(--g200)]" />
            <span className="text-[11px] text-[var(--tmuted)]">หรือพิมพ์เอง</span>
            <div className="flex-1 h-px bg-[var(--g200)]" />
          </div>

          <div className="flex gap-2">
            <input
              value={manual}
              onChange={(e) => setManual(e.target.value)}
              onKeyDown={(e) => { if (e.key === 'Enter') submitManual(); }}
              placeholder="พิมพ์ Serial หรือ Code..."
              className="flex-1 h-10 px-3 rounded-lg border border-[var(--g300)] text-[13px] focus:outline-none focus:border-[var(--blue)]"
            />
            <button
              onClick={submitManual}
              disabled={!manual.trim()}
              className="h-10 px-4 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold hover:bg-[var(--blue-d)] disabled:opacity-40"
            >
              ค้นหา
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}