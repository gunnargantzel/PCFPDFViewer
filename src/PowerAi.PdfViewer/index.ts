import { IInputs, IOutputs } from './generated/ManifestTypes';
import { getDocument, GlobalWorkerOptions } from 'pdfjs-dist';
// @ts-ignore
import pdfjsWorker from 'pdfjs-dist/build/pdf.worker.mjs';

GlobalWorkerOptions.workerSrc = pdfjsWorker as any;

type FitMode = 'auto'|'width'|'page';

export class PdfViewer implements ComponentFramework.StandardControl<IInputs, IOutputs> {

  private container!: HTMLDivElement;
  private ctx!: ComponentFramework.Context<IInputs>;
  private notifyOutputChanged!: () => void;
  private pdfData: ArrayBuffer | null = null;
  private resizeObserver?: ResizeObserver;
  private unhookPrint?: () => void;
  private toolbar?: HTMLDivElement;
  private canvas?: HTMLCanvasElement;
  private currentFit: FitMode = 'auto';

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this.ctx = context;
    this.notifyOutputChanged = notifyOutputChanged;
    this.container = container;
    this.container.classList.add('pai-pdf-wrap');

    if ('ResizeObserver' in window) {
      this.resizeObserver = new ResizeObserver(() => this.render());
      this.resizeObserver.observe(this.container);
    }

    this.applyPrintGuard(this.getBool(context.parameters.allowPrint, false));
  }

  public async updateView(context: ComponentFramework.Context<IInputs>): Promise<void> {
    this.ctx = context;
    this.currentFit = this.getFitMode(context.parameters.pageFit?.raw);
    const table = context.parameters.tableLogicalName?.raw || '';
    const recId = (context.parameters.recordId?.raw || '').replace(/[{}]/g, '');
    const col = context.parameters.fileColumnName?.raw || '';

    if (!table || !col || !recId) {
      this.renderMessage('PDF not configured (missing table/column/record).');
      return;
    }
    try {
      const buf = await this.fetchPdfBuffer(table, recId, col);
      this.pdfData = buf;
      await this.render();
    } catch (e:any) {
      this.renderMessage(`Failed to load PDF: ${e?.message || e}`);
    }
  }

  public getOutputs(): IOutputs { return {}; }
  public destroy(): void {
    this.resizeObserver?.disconnect();
    if (this.unhookPrint) this.unhookPrint();
  }

  private async render(): Promise<void> {
    if (!this.pdfData) return;
    const allowDownload = this.getBool(this.ctx.parameters.allowDownload, false);
    const allowPrint = this.getBool(this.ctx.parameters.allowPrint, false);
    const toolbarVisible = this.getBool(this.ctx.parameters.toolbarVisible, true);

    this.container.innerHTML = '';

    if (toolbarVisible) {
      this.toolbar = document.createElement('div');
      this.toolbar.className = 'pai-toolbar';

      const btnPrev = this.makeBtn('Prev', false);
      const btnNext = this.makeBtn('Next', false);
      const btnZoomIn = this.makeBtn('+', true, () => this.zoom(1.1));
      const btnZoomOut = this.makeBtn('−', true, () => this.zoom(1/1.1));
      const btnDownload = this.makeBtn('Download', allowDownload, () => this.download());
      const btnPrint = this.makeBtn('Print', allowPrint, () => window.print());

      this.toolbar.append(btnPrev, btnNext, btnZoomIn, btnZoomOut, btnDownload, btnPrint);
      this.container.appendChild(this.toolbar);
    }

    this.canvas = document.createElement('canvas');
    this.container.appendChild(this.canvas);

    const pdf = await getDocument({ data: new Uint8Array(this.pdfData) }).promise;
    const page = await pdf.getPage(1);

    const ctx2d = this.canvas.getContext('2d')!;
    const parentRect = this.container.getBoundingClientRect();
    const baseViewport = page.getViewport({ scale: 1 });
    const scale = this.calcScale(parentRect.width, parentRect.height, baseViewport.width, baseViewport.height, this.currentFit);
    const vp = page.getViewport({ scale });

    this.canvas.width = Math.floor(vp.width);
    this.canvas.height = Math.floor(vp.height);
    await page.render({ canvasContext: ctx2d, viewport: vp }).promise;
  }

  private zoom(factor:number) {
    if (!this.canvas) return;
    this.currentFit = 'width';
    const w = (this.canvas.offsetWidth || this.canvas.width) * factor;
    const h = (this.canvas.offsetHeight || this.canvas.height) * factor;
    this.canvas.style.width = `${w}px`;
    this.canvas.style.height = `${h}px`;
  }

  private download() {
    if (!this.pdfData) return;
    const blob = new Blob([this.pdfData], { type: 'application/pdf' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = 'document.pdf';
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
  }

  private makeBtn(label: string, enabled: boolean, onClick?: () => void): HTMLButtonElement {
    const btn = document.createElement('button');
    btn.textContent = label;
    btn.className = 'pai-btn';
    if (!enabled) btn.setAttribute('disabled', 'true');
    if (onClick) btn.addEventListener('click', onClick);
    return btn;
  }

  private applyPrintGuard(allowPrint: boolean) {
    if (this.unhookPrint) this.unhookPrint();
    const handler = (e: KeyboardEvent) => {
      if (!allowPrint && (e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'p') {
        e.preventDefault(); e.stopPropagation();
      }
    };
    window.addEventListener('keydown', handler, true);
    this.unhookPrint = () => window.removeEventListener('keydown', handler, true);
  }

  private async fetchPdfBuffer(table: string, id: string, col: string): Promise<ArrayBuffer> {
    const base = (this.ctx as any).page?.getClientUrl?.() || (this.ctx as any).page?.context?.getClientUrl?.();
    const url = `${base}/api/data/v9.2/${table}(${id})/${col}/$value`;
    const res = await fetch(url, { headers: { 'Accept': 'application/pdf' }, credentials: 'include' });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return await res.arrayBuffer();
  }

  private getBool(param: ComponentFramework.PropertyTypes.TwoOptionsProperty|undefined, def: boolean): boolean {
    if (!param || param.raw === null || param.raw === undefined) return def;
    return !!param.raw;
  }
  private getFitMode(raw: any): FitMode {
    if (raw === 1 || raw === 'width') return 'width';
    if (raw === 2 || raw === 'page') return 'page';
    return 'auto';
  }
  private calcScale(parentW:number, parentH:number, pageW:number, pageH:number, fit:FitMode): number {
    const scaleW = parentW / pageW;
    const scaleH = parentH / pageH;
    if (fit === 'width') return scaleW;
    if (fit === 'page') return Math.min(scaleW, scaleH);
    return Math.min(scaleW, scaleH);
  }

  private renderMessage(msg: string) {
    this.container.innerHTML = `<div style="color:#ccc;padding:.75rem;font-family:Segoe UI,Arial,sans-serif">${msg}</div>`;
  }
}
