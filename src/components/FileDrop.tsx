import { useId, useMemo, useState } from 'react';

type Props = {
  label: string;
  helperText?: string;
  value?: File | null;
  accept?: string;
  disabled?: boolean;
  onChange: (file: File | null) => void;
};

export function FileDrop({ label, helperText, value, accept = '.xlsx,.xls', disabled, onChange }: Props) {
  const inputId = useId();
  const [isDragging, setIsDragging] = useState(false);

  const fileName = useMemo(() => value?.name || '파일을 선택하거나 여기로 드래그하세요', [value]);

  return (
    <div className="field">
      <div className="fieldLabelRow">
        <div className="fieldLabel">{label}</div>
        {helperText ? <div className="fieldHelper">{helperText}</div> : null}
      </div>

      <label
        className={['fileDrop', isDragging ? 'fileDrop--drag' : '', disabled ? 'fileDrop--disabled' : '']
          .filter(Boolean)
          .join(' ')}
        htmlFor={inputId}
        onDragEnter={(e) => {
          e.preventDefault();
          if (disabled) return;
          setIsDragging(true);
        }}
        onDragOver={(e) => {
          e.preventDefault();
          if (disabled) return;
          setIsDragging(true);
        }}
        onDragLeave={(e) => {
          e.preventDefault();
          setIsDragging(false);
        }}
        onDrop={(e) => {
          e.preventDefault();
          setIsDragging(false);
          if (disabled) return;
          const f = e.dataTransfer.files?.[0] || null;
          onChange(f);
        }}
      >
        <input
          id={inputId}
          className="fileDropInput"
          type="file"
          accept={accept}
          disabled={disabled}
          onChange={(e) => onChange(e.target.files?.[0] || null)}
        />

        <div className="fileDropInner">
          <div className="fileDropTitle">{fileName}</div>
          <div className="fileDropMeta">{value ? `${Math.round(value.size / 1024).toLocaleString()} KB` : 'XLSX/XLS'}</div>
        </div>
      </label>

      {value ? (
        <div className="fieldActions">
          <button className="btn btn--secondary" type="button" onClick={() => onChange(null)} disabled={disabled}>
            파일 제거
          </button>
        </div>
      ) : null}
    </div>
  );
}

