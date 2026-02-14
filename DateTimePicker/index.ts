import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class DateTimePicker implements ComponentFramework.StandardControl<
  IInputs,
  IOutputs
> {
  private _container: HTMLDivElement;
  private _notifyOutputChanged: () => void;
  private _currentValue: Date | null = null;

  private _inputRow: HTMLDivElement;
  private _displayInput: HTMLInputElement;
  private _popup: HTMLDivElement;
  private _popupOpen = false;

  private _cur: Date = new Date();
  private _selectedDay: number | null = null;
  private _scrollOpen = false;
  private _is24h = true;
  private _ampm = "AM";

  private _calGrid: HTMLDivElement;
  private _monthYearLabel: HTMLSpanElement;
  private _headerArrow: HTMLSpanElement;
  private _calView: HTMLDivElement;
  private _scrollPanel: HTMLDivElement;
  private _monthList: HTMLDivElement;
  private _yearList: HTMLDivElement;
  private _hoursInput: HTMLInputElement;
  private _minutesInput: HTMLInputElement;
  private _ampmBtn: HTMLButtonElement;
  private _fmtLabel: HTMLSpanElement;
  private _tog: HTMLDivElement;
  private _prevBtn: HTMLButtonElement;
  private _nextBtn: HTMLButtonElement;

  private _MONTHS = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
  ];

  // eslint-disable-next-line @typescript-eslint/no-empty-function
  constructor() {}

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement,
  ): void {
    this._notifyOutputChanged = notifyOutputChanged;
    this._container = container;
    this._container.style.position = "relative";
    this._buildUI();
  }

  private _pad(n: number): string {
    return String(n).padStart(2, "0");
  }

  // FIX: helper to clamp an input's numeric value within its own min/max attributes
  private _clampInput(input: HTMLInputElement): void {
    const min = parseInt(input.min);
    const max = parseInt(input.max);
    let val = parseInt(input.value) || 0;
    if (!isNaN(min) && val < min) val = min;
    if (!isNaN(max) && val > max) val = max;
    input.value = this._pad(val);
  }

  private _buildUI(): void {
    // Input row
    this._inputRow = document.createElement("div");
    Object.assign(this._inputRow.style, {
      display: "flex",
      alignItems: "center",
      border: "1px solid #605e5c",
      borderRadius: "2px",
      background: "#ffffff",
      overflow: "hidden",
      cursor: "pointer",
      width: "100%",
    });

    this._displayInput = document.createElement("input");
    this._displayInput.type = "text";
    this._displayInput.placeholder = "DD/MM/YYYY HH:mm";
    this._displayInput.readOnly = true;
    Object.assign(this._displayInput.style, {
      border: "none",
      background: "transparent",
      fontSize: "14px",
      color: "#323130",
      padding: "6px 10px",
      flex: "1",
      outline: "none",
      cursor: "pointer",
    });

    const icon = document.createElement("span");
    icon.style.cssText =
      "padding:0 10px;flex-shrink:0;display:flex;align-items:center";
    icon.innerHTML = `<svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
      <rect x="1.5" y="3.5" width="13" height="11" rx="1.5" stroke="#0078d4" stroke-width="1.4"/>
      <path d="M1.5 7h13" stroke="#0078d4" stroke-width="1.4"/>
      <path d="M5 1.5v3M11 1.5v3" stroke="#0078d4" stroke-width="1.4" stroke-linecap="round"/>
      <circle cx="5.5" cy="10.5" r="0.8" fill="#0078d4"/>
      <circle cx="8" cy="10.5" r="0.8" fill="#0078d4"/>
      <circle cx="10.5" cy="10.5" r="0.8" fill="#0078d4"/>
    </svg>`;

    this._inputRow.appendChild(this._displayInput);
    this._inputRow.appendChild(icon);
    this._inputRow.addEventListener("click", () => this._togglePopup());

    // Popup
    this._popup = document.createElement("div");
    Object.assign(this._popup.style, {
      display: "none",
      position: "absolute",
      top: "100%",
      left: "0",
      background: "#ffffff",
      border: "1px solid #edebe9",
      borderRadius: "4px",
      width: "300px",
      zIndex: "9999",
      boxShadow: "0 4px 12px rgba(0,0,0,0.1)",
      marginTop: "2px",
    });

    this._popup.appendChild(this._buildHeader());
    this._calView = document.createElement("div");
    this._calGrid = document.createElement("div");
    this._calGrid.style.cssText =
      "display:grid;grid-template-columns:repeat(7,1fr);padding:8px 8px 4px";
    this._calView.appendChild(this._calGrid);
    this._popup.appendChild(this._calView);

    this._scrollPanel = this._buildScrollPanel();
    this._popup.appendChild(this._scrollPanel);
    this._popup.appendChild(this._buildTimeSection());
    this._popup.appendChild(this._buildFormatToggle());
    this._popup.appendChild(this._buildFooter());

    this._container.appendChild(this._inputRow);
    this._container.appendChild(this._popup);

    this._renderCal();
  }

  private _buildHeader(): HTMLDivElement {
    const header = document.createElement("div");
    header.style.cssText =
      "display:flex;align-items:center;justify-content:space-between;padding:10px 12px;border-bottom:1px solid #edebe9";

    this._prevBtn = document.createElement("button");
    this._prevBtn.textContent = "‹";
    this._styleNavBtn(this._prevBtn);
    this._prevBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      if (!this._scrollOpen) {
        this._cur.setMonth(this._cur.getMonth() - 1);
        this._renderCal();
      }
    });

    const centerBtn = document.createElement("button");
    Object.assign(centerBtn.style, {
      background: "none",
      border: "none",
      cursor: "pointer",
      fontSize: "14px",
      fontWeight: "500",
      color: "#323130",
      padding: "4px 8px",
      borderRadius: "4px",
      display: "flex",
      alignItems: "center",
      gap: "4px",
    });

    this._monthYearLabel = document.createElement("span");
    this._headerArrow = document.createElement("span");
    this._headerArrow.style.cssText = "font-size:10px;color:#605e5c";
    this._headerArrow.textContent = "▼";

    centerBtn.appendChild(this._monthYearLabel);
    centerBtn.appendChild(this._headerArrow);
    centerBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      this._toggleScroll();
    });

    this._nextBtn = document.createElement("button");
    this._nextBtn.textContent = "›";
    this._styleNavBtn(this._nextBtn);
    this._nextBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      if (!this._scrollOpen) {
        this._cur.setMonth(this._cur.getMonth() + 1);
        this._renderCal();
      }
    });

    header.appendChild(this._prevBtn);
    header.appendChild(centerBtn);
    header.appendChild(this._nextBtn);
    return header;
  }

  private _styleNavBtn(btn: HTMLButtonElement): void {
    Object.assign(btn.style, {
      background: "none",
      border: "none",
      cursor: "pointer",
      color: "#605e5c",
      fontSize: "18px",
      padding: "2px 8px",
      borderRadius: "4px",
      lineHeight: "1",
    });
  }

  private _buildScrollPanel(): HTMLDivElement {
    const panel = document.createElement("div");
    panel.style.cssText = "display:none;padding:8px;gap:8px";

    const monthCol = this._buildScrollCol("Month");
    this._monthList = monthCol.list;

    const yearCol = this._buildScrollCol("Year");
    this._yearList = yearCol.list;

    panel.appendChild(monthCol.col);
    panel.appendChild(yearCol.col);
    return panel;
  }

  private _buildScrollCol(label: string): {
    col: HTMLDivElement;
    list: HTMLDivElement;
  } {
    const col = document.createElement("div");
    col.style.cssText = "flex:1;display:flex;flex-direction:column;gap:2px";

    const lbl = document.createElement("div");
    lbl.style.cssText =
      "font-size:11px;color:#605e5c;text-align:center;padding:4px 0;border-bottom:1px solid #edebe9;margin-bottom:4px";
    lbl.textContent = label;

    const list = document.createElement("div");
    list.style.cssText =
      "height:180px;overflow-y:auto;display:flex;flex-direction:column;gap:1px";

    col.appendChild(lbl);
    col.appendChild(list);
    return { col, list };
  }

  private _buildScrollLists(): void {
    this._monthList.innerHTML = "";
    this._yearList.innerHTML = "";

    this._MONTHS.forEach((m, i) => {
      const el = document.createElement("div");
      el.style.cssText =
        "text-align:center;padding:6px 4px;font-size:13px;cursor:pointer;border-radius:4px;color:#323130";
      if (i === this._cur.getMonth()) {
        el.style.background = "#0078d4";
        el.style.color = "#fff";
      }
      el.textContent = m;
      el.addEventListener("click", () => {
        this._cur.setMonth(i);
        this._buildScrollLists();
        this._updateHeader();
      });
      this._monthList.appendChild(el);
    });

    for (let y = 1970; y <= 2050; y++) {
      const el = document.createElement("div");
      el.style.cssText =
        "text-align:center;padding:6px 4px;font-size:13px;cursor:pointer;border-radius:4px;color:#323130";
      if (y === this._cur.getFullYear()) {
        el.style.background = "#0078d4";
        el.style.color = "#fff";
      }
      el.textContent = String(y);
      el.addEventListener("click", () => {
        this._cur.setFullYear(y);
        this._buildScrollLists();
        this._updateHeader();
      });
      this._yearList.appendChild(el);
    }

    setTimeout(() => {
      const selM = this._monthList.querySelector(
        'div[style*="0078d4"]',
      ) as HTMLElement;
      if (selM)
        this._monthList.scrollTop =
          selM.offsetTop -
          this._monthList.offsetHeight / 2 +
          selM.offsetHeight / 2;
      const selY = this._yearList.querySelector(
        'div[style*="0078d4"]',
      ) as HTMLElement;
      if (selY)
        this._yearList.scrollTop =
          selY.offsetTop -
          this._yearList.offsetHeight / 2 +
          selY.offsetHeight / 2;
    }, 0);
  }

  private _buildTimeSection(): HTMLDivElement {
    const section = document.createElement("div");
    section.style.cssText = "padding:10px 12px;border-top:1px solid #edebe9";

    const label = document.createElement("div");
    label.style.cssText = "font-size:12px;color:#605e5c;margin-bottom:8px";
    label.textContent = "Time";

    const controls = document.createElement("div");
    controls.style.cssText =
      "display:flex;align-items:center;justify-content:center;gap:8px";

    const timeInputs = document.createElement("div");
    timeInputs.style.cssText = "display:flex;align-items:center;gap:4px";

    this._hoursInput = document.createElement("input");
    this._hoursInput.type = "number";
    this._hoursInput.value = "00";
    // FIX: define min/max for 24h mode (default)
    this._hoursInput.min = "0";
    this._hoursInput.max = "23";
    this._styleTimeInput(this._hoursInput);
    // FIX: clamp on every keystroke and on blur
    this._hoursInput.addEventListener("input", () =>
      this._clampInput(this._hoursInput),
    );
    this._hoursInput.addEventListener("blur", () =>
      this._clampInput(this._hoursInput),
    );

    const sep = document.createElement("span");
    sep.textContent = ":";
    sep.style.cssText = "font-size:16px;color:#605e5c;font-weight:500";

    this._minutesInput = document.createElement("input");
    this._minutesInput.type = "number";
    this._minutesInput.min = "0";
    this._minutesInput.max = "59";
    this._minutesInput.value = "00";
    this._styleTimeInput(this._minutesInput);
    // FIX: clamp on every keystroke and on blur
    this._minutesInput.addEventListener("input", () =>
      this._clampInput(this._minutesInput),
    );
    this._minutesInput.addEventListener("blur", () =>
      this._clampInput(this._minutesInput),
    );

    this._ampmBtn = document.createElement("button");
    this._ampmBtn.textContent = "AM";
    this._ampmBtn.style.cssText =
      "display:none;align-items:center;justify-content:center;width:44px;height:34px;font-size:13px;font-weight:500;border:1px solid #0078d4;background:#0078d4;color:#fff;cursor:pointer;border-radius:4px;margin-left:4px;padding:0";
    this._ampmBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      this._ampm = this._ampm === "AM" ? "PM" : "AM";
      this._ampmBtn.textContent = this._ampm;
    });

    timeInputs.appendChild(this._hoursInput);
    timeInputs.appendChild(sep);
    timeInputs.appendChild(this._minutesInput);
    controls.appendChild(timeInputs);
    controls.appendChild(this._ampmBtn);
    section.appendChild(label);
    section.appendChild(controls);
    return section;
  }

  private _styleTimeInput(input: HTMLInputElement): void {
    Object.assign(input.style, {
      width: "46px",
      textAlign: "center",
      border: "1px solid #605e5c",
      borderRadius: "4px",
      padding: "5px 4px",
      fontSize: "15px",
      color: "#323130",
      background: "#ffffff",
      outline: "none",
    });
  }

  private _buildFormatToggle(): HTMLDivElement {
    const row = document.createElement("div");
    row.style.cssText =
      "display:flex;align-items:center;gap:6px;padding:8px 12px;border-top:1px solid #edebe9;background:#ffffff;cursor:pointer;user-select:none";

    this._tog = document.createElement("div");
    this._tog.style.cssText =
      "position:relative;width:32px;height:18px;background:#0078d4;border-radius:9px;flex-shrink:0;transition:background .15s";

    const thumb = document.createElement("div");
    thumb.style.cssText =
      "position:absolute;top:2px;left:14px;width:14px;height:14px;background:#fff;border-radius:50%;transition:transform .15s";
    this._tog.appendChild(thumb);

    this._fmtLabel = document.createElement("span");
    this._fmtLabel.style.cssText = "font-size:12px;color:#605e5c";
    this._fmtLabel.textContent = "24h format";

    row.appendChild(this._tog);
    row.appendChild(this._fmtLabel);
    row.addEventListener("click", () => this._switchFormat());
    return row;
  }

  private _buildFooter(): HTMLDivElement {
    const footer = document.createElement("div");
    footer.style.cssText =
      "display:flex;justify-content:flex-end;gap:8px;padding:8px 12px;border-top:1px solid #edebe9";

    const clearBtn = document.createElement("button");
    clearBtn.textContent = "Clear";
    this._styleBtn(clearBtn, false);
    clearBtn.addEventListener("click", () => this._clearValue());

    const confirmBtn = document.createElement("button");
    confirmBtn.textContent = "Confirm";
    this._styleBtn(confirmBtn, true);
    confirmBtn.addEventListener("click", () => {
      if (this._scrollOpen) {
        this._toggleScroll();
        return;
      }
      this._confirmValue();
    });

    footer.appendChild(clearBtn);
    footer.appendChild(confirmBtn);
    return footer;
  }

  private _styleBtn(btn: HTMLButtonElement, primary: boolean): void {
    btn.style.cssText = primary
      ? "font-size:13px;padding:5px 16px;border-radius:4px;cursor:pointer;border:1px solid #0078d4;background:#0078d4;color:#fff"
      : "font-size:13px;padding:5px 16px;border-radius:4px;cursor:pointer;border:1px solid #605e5c;background:transparent;color:#323130";
  }

  private _updateHeader(): void {
    this._monthYearLabel.textContent =
      this._MONTHS[this._cur.getMonth()] + " " + this._cur.getFullYear();
    this._headerArrow.textContent = this._scrollOpen ? "▲" : "▼";
  }

  private _renderCal(): void {
    this._updateHeader();
    const days = ["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"];
    let html = days
      .map(
        (d) =>
          `<div style="text-align:center;font-size:11px;color:#605e5c;padding:4px 0">${d}</div>`,
      )
      .join("");
    const first = new Date(
      this._cur.getFullYear(),
      this._cur.getMonth(),
      1,
    ).getDay();
    const total = new Date(
      this._cur.getFullYear(),
      this._cur.getMonth() + 1,
      0,
    ).getDate();
    const prevTotal = new Date(
      this._cur.getFullYear(),
      this._cur.getMonth(),
      0,
    ).getDate();
    for (let i = 0; i < first; i++)
      html += `<div style="text-align:center;font-size:13px;color:#c8c6c4;padding:5px 0;border-radius:50%">${prevTotal - first + 1 + i}</div>`;
    for (let d = 1; d <= total; d++) {
      const sel = d === this._selectedDay;
      html += `<div data-day="${d}" style="text-align:center;font-size:13px;padding:5px 0;border-radius:50%;cursor:pointer;color:${sel ? "#fff" : "#323130"};background:${sel ? "#0078d4" : "transparent"}">${d}</div>`;
    }
    const rem = (first + total) % 7 === 0 ? 0 : 7 - ((first + total) % 7);
    for (let j = 1; j <= rem; j++)
      html += `<div style="text-align:center;font-size:13px;color:#c8c6c4;padding:5px 0">${j}</div>`;
    this._calGrid.innerHTML = html;
    this._calGrid.querySelectorAll("[data-day]").forEach((el) => {
      el.addEventListener("click", () => {
        this._selectedDay = parseInt(
          (el as HTMLElement).getAttribute("data-day")!,
        );
        this._renderCal();
      });
      (el as HTMLElement).addEventListener("mouseover", () => {
        if (
          parseInt((el as HTMLElement).getAttribute("data-day")!) !==
          this._selectedDay
        )
          (el as HTMLElement).style.background = "#f3f2f1";
      });
      (el as HTMLElement).addEventListener("mouseout", () => {
        if (
          parseInt((el as HTMLElement).getAttribute("data-day")!) !==
          this._selectedDay
        )
          (el as HTMLElement).style.background = "transparent";
      });
    });
  }

  private _toggleScroll(): void {
    this._scrollOpen = !this._scrollOpen;
    this._scrollPanel.style.display = this._scrollOpen ? "flex" : "none";
    this._calView.style.display = this._scrollOpen ? "none" : "block";
    this._prevBtn.style.visibility = this._scrollOpen ? "hidden" : "visible";
    this._nextBtn.style.visibility = this._scrollOpen ? "hidden" : "visible";
    if (this._scrollOpen) this._buildScrollLists();
    else this._renderCal();
    this._updateHeader();
  }

  private _togglePopup(): void {
    this._popupOpen = !this._popupOpen;
    this._popup.style.display = this._popupOpen ? "block" : "none";
  }

  private _get24h(): number {
    const h = parseInt(this._hoursInput.value) || 0;
    if (this._is24h) return h;
    if (this._ampm === "AM") return h === 12 ? 0 : h;
    return h === 12 ? 12 : h + 12;
  }

  private _switchFormat(): void {
    const h24 = this._get24h();
    this._is24h = !this._is24h;
    const thumb = this._tog.querySelector("div") as HTMLElement;
    if (this._is24h) {
      // FIX: restore 24h bounds
      this._hoursInput.min = "0";
      this._hoursInput.max = "23";
      this._hoursInput.value = this._pad(h24);
      this._ampmBtn.style.display = "none";
      this._tog.style.background = "#0078d4";
      thumb.style.left = "14px";
      this._fmtLabel.textContent = "24h format";
    } else {
      // FIX: 12h mode — valid range is 1..12
      this._hoursInput.min = "1";
      this._hoursInput.setAttribute("max", "12");
      if (h24 === 0) {
        this._hoursInput.value = "12";
        this._ampm = "AM";
      } else if (h24 < 12) {
        this._hoursInput.value = this._pad(h24);
        this._ampm = "AM";
      } else if (h24 === 12) {
        this._hoursInput.value = "12";
        this._ampm = "PM";
      } else {
        this._hoursInput.value = this._pad(h24 - 12);
        this._ampm = "PM";
      }
      this._ampmBtn.textContent = this._ampm;
      this._ampmBtn.style.display = "inline-flex";
      this._tog.style.background = "#bbb";
      thumb.style.left = "2px";
      this._fmtLabel.textContent = "12h format (AM/PM)";
    }
  }

  private _confirmValue(): void {
    if (!this._selectedDay) return;
    const h24 = this._get24h();
    const m = parseInt(this._minutesInput.value) || 0;
    const date = new Date(
      this._cur.getFullYear(),
      this._cur.getMonth(),
      this._selectedDay,
      h24,
      m,
      0,
    );
    this._currentValue = date;
    const day = this._pad(this._selectedDay);
    const month = this._pad(this._cur.getMonth() + 1);
    const year = this._cur.getFullYear();
    let timeStr: string;
    if (this._is24h) {
      timeStr = `${this._pad(h24)}:${this._pad(m)}`;
    } else {
      const display = h24 === 0 || h24 === 12 ? 12 : h24 % 12;
      timeStr = `${this._pad(display)}:${this._pad(m)} ${this._ampm}`;
    }
    this._displayInput.value = `${day}/${month}/${year} ${timeStr}`;
    this._popup.style.display = "none";
    this._popupOpen = false;
    this._notifyOutputChanged();
  }

  private _clearValue(): void {
    this._currentValue = null;
    this._selectedDay = null;
    this._displayInput.value = "";
    this._hoursInput.value = "00";
    this._minutesInput.value = "00";
    if (this._scrollOpen) this._toggleScroll();
    else this._renderCal();
    this._notifyOutputChanged();
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    const raw = context.parameters.dateTimeValue.raw;
    if (raw && !this._popupOpen) {
      this._currentValue = raw;
      this._cur = new Date(raw);
      this._selectedDay = raw.getDate();
      const h = raw.getHours();
      const m = raw.getMinutes();
      if (this._is24h) {
        this._hoursInput.value = this._pad(h);
      } else {
        this._ampm = h >= 12 ? "PM" : "AM";
        this._hoursInput.value = this._pad(h === 0 ? 12 : h > 12 ? h - 12 : h);
        this._ampmBtn.textContent = this._ampm;
      }
      this._minutesInput.value = this._pad(m);
      const month = this._pad(raw.getMonth() + 1);
      const day = this._pad(raw.getDate());
      const year = raw.getFullYear();
      let timeStr: string;
      if (this._is24h) {
        timeStr = `${this._pad(h)}:${this._pad(m)}`;
      } else {
        const display = h === 0 || h === 12 ? 12 : h % 12;
        timeStr = `${this._pad(display)}:${this._pad(m)} ${this._ampm}`;
      }
      this._displayInput.value = `${day}/${month}/${year} ${timeStr}`;
      this._renderCal();
    }
  }

  public getOutputs(): IOutputs {
    return {
      dateTimeValue: this._currentValue ?? undefined,
    };
  }

  public destroy(): void {
    this._inputRow.removeEventListener("click", () => this._togglePopup());
  }
}
