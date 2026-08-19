import CaseGallery from "./case-gallery";
import EvaluationExplorer from "./evaluation-explorer";

const formats = [
  {
    key: "PPT",
    title: "PowerPoint",
    status: "完成",
    tone: "gold",
    image: "/ppt-reward-sample.png",
    width: 1280,
    height: 720,
    summary: "逐 slide 截图、文本、结构信号与失败覆盖已完成。",
    metrics: ["真实 8-slide deck", "1280×720 render", "部分失败可恢复"],
  },
  {
    key: "DOC",
    title: "Word",
    status: "完成",
    tone: "blue",
    image: "/word-reward-sample.png",
    width: 794,
    height: 1123,
    summary: "稳定分页走 page unit，其余文档按标题 section 或段落窗口评分。",
    metrics: ["标题 / 表格映射", "同源 HTML 截图", "OOXML 预检"],
  },
  {
    key: "XLS",
    title: "Excel",
    status: "完成",
    tone: "green",
    image: "/excel-reward-sample.png",
    width: 1600,
    height: 1200,
    summary: "逐 sheet DOM 激活、稀疏单元格、共享公式与视觉 artifact 已完成。",
    metrics: ["hidden sheet 对齐", "chart crop", "20 shapes smoke"],
  },
];

const pipeline = [
  ["01", "Office schema / runner", "done"],
  ["02", "格式 rubric / 多图输入", "done"],
  ["03", "PPTX adapter", "done"],
  ["04", "DOCX adapter", "done"],
  ["05", "XLSX adapter", "done"],
  ["06", "文档 reward 聚合", "done"],
  ["07", "score-office CLI", "done"],
  ["08", "端到端验证 / 文档", "done"],
];

const testHistory = [
  ["基线", 135],
  ["Runner", 208],
  ["Prompt contract", 284],
  ["PPTX", 345],
  ["DOCX", 487],
  ["XLSX", 595],
  ["Reward engine", 638],
  ["Office CLI", 666],
  ["Final", 678],
];

const evidence = [
  {
    format: "PPT",
    result: "直接 .pptx 输入",
    proof: "逐页文本、截图、issues、validate",
    state: "通过",
  },
  {
    format: "Word",
    result: "标题 + 表格 + 子节",
    proof: "section 非重叠，表格单元格进入上下文",
    state: "通过",
  },
  {
    format: "Excel",
    result: "hidden / chart-only / shape-only",
    proof: "HTML tab 身份、chart crop、20 shapes",
    state: "通过",
  },
  {
    format: "Core",
    result: "安全与资源边界",
    proof: "ZIP 预检、进程树终止、输出硬上限",
    state: "通过",
  },
  {
    format: "Reward",
    result: "partial / failed / strict",
    proof: "维度 coverage、校准隔离、issue 降级",
    state: "通过",
  },
  {
    format: "CLI",
    result: "manifest / replay / 并发",
    proof: "单次 probe、输入保序、敏感错误脱敏",
    state: "通过",
  },
  {
    format: "Final",
    result: "真实三格式 + 全量回归",
    proof: "678 tests、双重复审、fast-forward main",
    state: "通过",
  },
];

export default function Home() {
  return (
    <main>
      <header className="topbar">
        <a className="brand" href="#status" aria-label="Office Reward Build Log">
          <span className="brand-mark">OR</span>
          <span>Office Reward Build Log</span>
        </a>
        <nav aria-label="页面导航">
          <a href="/rubric">人工标注</a>
          <a href="#formats">格式</a>
          <a href="#cases">Cases</a>
          <a href="#mechanism">机制</a>
          <a href="#prompts">Prompt</a>
          <a href="#pipeline">流水线</a>
          <a href="#evidence">验证</a>
        </nav>
        <div className="asof">2026.08.05 · CASES</div>
      </header>

      <section className="status-band" id="status">
        <div className="status-copy">
          <p className="eyebrow">BUILD COMPLETE / MERGED TO MAIN</p>
          <h1>Office reward 全链路完成</h1>
          <p className="lede">
            PPT、Word、Excel 的 adapter、reward engine、`score-office`、真实 replay
            集成与文档均已完成。最终修复通过双重复审，并已合并 main。
          </p>
        </div>
        <div className="status-meter" aria-label="总体完成度 100%">
          <div className="meter-header">
            <span>总体完成度</span>
            <strong>100%</strong>
          </div>
          <div className="meter-track">
            <span style={{ width: "100%" }} />
          </div>
          <div className="meter-legend">
            <span><i className="dot done" />8 完成</span>
            <span><i className="dot active" />0 进行中</span>
            <span><i className="dot queued" />0 排队</span>
          </div>
        </div>
        <dl className="headline-stats">
          <div>
            <dt>完整测试</dt>
            <dd>678</dd>
          </div>
          <div>
            <dt>真实格式</dt>
            <dd>3</dd>
          </div>
          <div>
            <dt>当前分支</dt>
            <dd className="branch">main</dd>
          </div>
        </dl>
      </section>

      <section className="section" id="formats">
        <div className="section-heading">
          <div>
            <p className="eyebrow">FORMAT TRACKS</p>
            <h2>格式进展</h2>
          </div>
          <p>每条链路都使用真实 OfficeCLI 输出，不以静态 mock 代替最终证据。</p>
        </div>

        <div className="format-grid">
          {formats.map((format) => (
            <article className={`format-item ${format.tone}`} key={format.key}>
              <div className="format-media">
                <img
                  src={format.image}
                  width={format.width}
                  height={format.height}
                  alt={`${format.title} reward 渲染样例`}
                  loading="lazy"
                  decoding="async"
                />
              </div>
              <div className="format-body">
                <div className="format-title">
                  <span className="format-key">{format.key}</span>
                  <div>
                    <h3>{format.title}</h3>
                    <span className="status-label">{format.status}</span>
                  </div>
                </div>
                <p>{format.summary}</p>
                <ul>
                  {format.metrics.map((metric) => <li key={metric}>{metric}</li>)}
                </ul>
              </div>
            </article>
          ))}
        </div>
      </section>

      <CaseGallery />

      <EvaluationExplorer />

      <section className="section pipeline-section" id="pipeline">
        <div className="section-heading">
          <div>
            <p className="eyebrow">DELIVERY PIPELINE</p>
            <h2>八步实现链路</h2>
          </div>
          <p>所有任务经过测试先行、规格审查和代码质量审查。</p>
        </div>
        <ol className="pipeline-list">
          {pipeline.map(([number, label, state]) => (
            <li className={state} key={number}>
              <span className="step-number">{number}</span>
              <span className="step-label">{label}</span>
              <span className="step-state">
                {state === "done" ? "DONE" : state === "active" ? "HARDENING" : "QUEUED"}
              </span>
            </li>
          ))}
        </ol>
      </section>

      <section className="section test-section">
        <div className="section-heading">
          <div>
            <p className="eyebrow">REGRESSION COVERAGE</p>
            <h2>测试增长</h2>
          </div>
          <p>从 PPT 单格式基线增长到三格式适配与共享安全边界。</p>
        </div>
        <div className="test-layout">
          <div className="test-chart" aria-label="测试数量增长图">
            {testHistory.map(([label, value]) => {
              const amount = Number(value);
              return (
                <div className="test-row" key={label}>
                  <span>{label}</span>
                  <div className="bar-track">
                    <i style={{ width: `${Math.round((amount / 678) * 100)}%` }} />
                  </div>
                  <strong>{amount}</strong>
                </div>
              );
            })}
          </div>
          <div className="review-log">
            <p className="review-title">审查中修复的关键问题</p>
            <ul>
              <li>权威 schema CRC 与真实非 resident 输出</li>
              <li>Word 标题、表格、分页与内容控件映射</li>
              <li>Excel hidden sheet、稀疏单元格与视觉 artifact</li>
              <li>ZIP 炸弹、输出上限、进程树与路径泄漏</li>
              <li>strict 缺失维度与 issue-summary 非关键降级</li>
              <li>CLI 单次 probe、未知类型与 manifest 错误脱敏</li>
              <li>CSP/外网隔离、used-range、像素与 unit 硬上限</li>
            </ul>
          </div>
        </div>
      </section>

      <section className="section evidence-section" id="evidence">
        <div className="section-heading">
          <div>
            <p className="eyebrow">VERIFIED EVIDENCE</p>
            <h2>真实验证记录</h2>
          </div>
          <p>smoke 结果只证明工程链路可运行，不替代后续人类校准。</p>
        </div>
        <div className="evidence-table" role="table" aria-label="真实验证记录">
          <div className="evidence-row evidence-head" role="row">
            <span>格式</span><span>场景</span><span>证据</span><span>状态</span>
          </div>
          {evidence.map((item) => (
            <div className="evidence-row" role="row" key={`${item.format}-${item.result}`}>
              <strong>{item.format}</strong>
              <span>{item.result}</span>
              <span>{item.proof}</span>
              <span className="pass">{item.state}</span>
            </div>
          ))}
        </div>
      </section>

      <section className="next-band">
        <div>
          <p className="eyebrow">DELIVERY STATUS</p>
          <h2>已合并 main</h2>
        </div>
        <div className="next-items">
          <span>PPT / Word / Excel</span>
          <span>score-office CLI</span>
          <span>678 tests</span>
          <span>真实 OfficeCLI</span>
          <span>规格与质量批准</span>
        </div>
      </section>

      <footer>
        <span>Office Reward / autonomous build</span>
        <span>Last updated · August 5, 2026</span>
      </footer>
    </main>
  );
}
