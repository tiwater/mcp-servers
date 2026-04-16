using Dockit.Docx;
using Xunit;

namespace Dockit.Docx.Tests;

public class ResolverTests
{
    [Fact]
    public void Resolve_emits_the_first_ana03_slice_from_plan_and_source_exports()
    {
        var plan = CreateAna03Plan();
        var stabilityDataPath = WriteTempJson("""
        [
          {
            "sheet": "Sheet1",
            "rows": [
              ["名称", "", "", "", "", "HSPXXXXDS", "", "", "", "", "", "", "", "", ""],
              ["批号", "", "", "", "", "250810S", "", "", "", "", "", "", "", "", ""],
              ["颜色", "", "与黄色4号标准比色液比较，不得更深", "Not more intensely than EP Y4 reference solution", "浅于黄色0.5号标准比色液\nLess colored than EP Y7 reference solution", "浅于黄色0.5号标准比色液\nLess colored than EP Y7 reference solution", "浅于黄色0.5号标准比色液\nLess colored than EP Y7 reference solution", "浅于黄色0.5号标准比色液；\nLess colored than EPY7 reference solution.", "浅于黄色2号标准比色液\nLess colored than EP Y6 reference solution", "浅于黄色2号标准比色液\nLess colored than EP Y5 reference solution", "浅于黄色0.5号标准比色液\nLess colored than EP Y7 reference solution", "浅于黄色0.5号标准比色液\nLess colored than EP Y7 reference solution", "浅于黄色0.5号标准比色液\nLess colored than EP Y7 reference solution", "浅于黄色0.5号标准比色液\nLess colored than EP Y7 reference solution", "浅于黄色0.5号标准比色液\nLess colored than EP Y7 reference solution"],
              ["澄清度", "", "与4号浊度标准液比较，不得更浓", "Not more tuebid than reference suspension IV", "低于0.5号浊度标准液\nLower than reference suspension Ⅰ.", "低于1号浊度标准液\nLower than reference suspension Ⅰ.", "低于1号浊度标准液\nLower than reference suspension Ⅰ.", "低于1号浊度标准液\nLower than reference suspension Ⅰ.", "低于1号浊度标准液\nLower than reference suspension Ⅰ", "低于1号浊度标准液\nLower than reference suspension Ⅰ", "低于1号浊度标准液\nLower than reference suspension Ⅰ", "低于1号浊度标准液\nLower than reference suspension Ⅰ", "低于1号浊度标准液\nLower than reference suspension Ⅰ", "低于1号浊度标准液\nLower than reference suspension Ⅰ", "低于1号浊度标准液\nLower than reference suspension Ⅰ."],
              ["PH", "", "5.2-6.0", "5.2-6.0", "5.6", "5.6", "5.6", "5.6", "5.5", "5.5", "5.6", "5.6", "5.6", "5.6", "5.6"],
              ["蛋白质含量", "", "9.0-11.0\nmg/ml", "9.0-11.0\nmg/mL", "10.3 \nmg/ml", "10.2", "10.4", "10.2", "10.2", "10.2", "10.3", "10.3", "10.3", "10.3", "10.2"]
            ]
          }
        ]
        """);
        var qualityStandardPath = WriteTempJson("""
        [
          {
            "type": "table",
            "rows": [
              ["质量属性", "测试项", "接受标准", "检测方法"],
              ["理化检定", "颜色", "与黄色4号标准比色液比较，不得更深", "中国药典：2025年版四部通则0901 溶液颜色检查法SOP03-5-2022"],
              ["理化检定", "澄清度", "与4号浊度标准液比较，不得更浓", "中国药典：2025年版四部通则0902 澄清度检查法SOP03-5-2022"],
              ["理化检定", "pH", "5.2-6.0", "中国药典：2025年版四部通则0631 pH值测定法SOP03-5-2021"],
              ["含量", "蛋白质含量", "9.0-11.0\nmg/ml", "SOP03-5-2028"]
            ]
          }
        ]
        """);
        var protocolPath = WriteTempJson("""
        [
          {
            "type": "table",
            "rows": [
              ["名称Name", "HSPXXXXDS"],
              ["批号Batch No.", "YYYY"],
              ["分装规格Strength", "3 ml/瓶3 mL/bottle"],
              ["生产商Manufacturer", "杭州皓阳生物技术有限公司Hangzhou HealSun Biopharm Co., Ltd."]
            ]
          }
        ]
        """);
        var inspectionReportPath = WriteTempJson("""
        [
          {
            "type": "table",
            "rows": [
              ["产品名称Product name", "HSPXXXXDS原液"],
              ["目标浓度 Target Concentration", "10.0 mg/ml", "项 目 号Project No.", "HSPXXXX"],
              ["批号Batch No.", "250810S", "生产日期Production Date", "2025-09-08"],
              ["总数量Total Quantity", "37.8 g", "有效期Expiration Date", "2028-09-07"],
              ["贮存条件Storage Condition", "≤-60 ℃，避光保存", "报告日期Report Date", "2026-02-03"],
              ["质量标准Specification No.", "XXXX-STP03-1-001（01）"],
              ["生产厂家Manufacturer", "杭州皓阳生物技术有限公司"]
            ]
          }
        ]
        """);
        var reportPath = WriteTempJson("""
        [
          { "type": "table", "rows": [["A"]] },
          { "type": "table", "rows": [["B"]] },
          {
            "type": "table",
            "rows": [
              ["条件Condition", "取样点Sampling Point", "放入时间Time to Place the Sample", "计划取样时间Planned Sampling Time", "实际取样时间Actual Sampling Time"],
              ["", "", "", "", ""],
              ["高温试验", "1 M", "", "", ""],
              ["", "2 M", "", "", ""],
              ["", "3 M", "", "", ""],
              ["光照试验", "日光＋紫外", "", "", ""],
              ["", "13 D", "", "", ""],
              ["", "光照试验阴性对照", "", "", ""],
              ["", "13 D", "", "", ""],
              ["冻融试验", "3 C", "", "", ""],
              ["", "5 C", "", "", ""]
            ]
          },
          { "type": "table", "rows": [["D"]] },
          {
            "type": "table",
            "rows": [
              ["产品名称Product Name", "placeholder"],
              ["生产日期Date of Manufacture", "placeholder"],
              ["生产厂家Manufacturer", "placeholder"],
              ["批量Batch Size", "placeholder"],
              ["批号Batch No.", "placeholder"],
              ["规格Strength", "placeholder"]
            ]
          },
          { "type": "table", "rows": [["F"]] },
          { "type": "table", "rows": [["G"]] },
          { "type": "table", "rows": [["H"]] },
          { "type": "table", "rows": [["I"]] },
          {
            "type": "table",
            "rows": [
              ["检项Test item", "接受标准Acceptance criteria", "T0", "反复冻融试验（≤ -60 ℃—室温）Repeated freeze-thaw testing(≤ -60 ℃—room temperature)"],
              ["", "", "", "3 cycles"],
              ["颜色", "", "", "", ""],
              ["", "", "", "", ""],
              ["澄清度", "", "", "", ""],
              ["", "", "", "", ""],
              ["PH", "", "", "", ""],
              ["蛋白质含量", "", "", "", ""]
            ]
          }
        ]
        """);
        var samplingPlanPath = WriteTempJson("""
        [
          {
            "sheet": "Sheet1",
            "rows": [
              ["放样日期", "名称", "批号", "考察条件", "考察周期", "取样数量（支）", "到期日", "实际取样日期"],
              ["2025.09.23", "HSPXXXXDS", "250810S", "高温（25℃±2℃/60%RH±5%RH）避光", "1个月", "检3/检备2", "2025.10.23（周四", "2025-10-23"],
              ["2025.09.23", "HSPXXXXDS", "250810S", "高温（25℃±2℃/60%RH±5%RH）避光", "2个月", "检3/检备2", "2025.11.23（周日", "2025-11-23"],
              ["2025.09.23", "HSPXXXXDS", "250810S", "高温（25℃±2℃/60%RH±5%RH）避光", "3个月", "检3/检备2", "2025.12.23（周二", "2025-12-23"],
              ["2025.09.26", "HSPXXXXDS", "250810S", "光照（日光＋紫外/25℃±2℃/60%RH±5%RH）", "6天", "检3/检备2", "2025.10.02（休", "2025-10-02"],
              ["2025.09.26", "HSPXXXXDS", "250810S", "光照（日光＋紫外/25℃±2℃/60%RH±5%RH）", "13天", "检3/检备2", "2025.10.09（周四", "2025-10-09"],
              ["2025.09.26", "HSPXXXXDS", "250810S", "光照阴性对照（25℃±2℃）", "6天", "检3/检备2", "2025.10.02（休", "2025-10-02"],
              ["2025.09.26", "HSPXXXXDS", "250810S", "光照阴性对照（25℃±2℃）", "13天", "检3/检备2", "2025.10.09（周四", "2025-10-09"],
              ["2025.09.23", "HSPXXXXDS", "250810S", "冻融试验（≤-60℃—室温冻融）", "3次", "检3/检备2", "2025.10.09（周四", "2025-10-09"],
              ["2025.09.23", "HSPXXXXDS", "250810S", "冻融试验（≤-60℃—室温冻融）", "5次", "检3/检备2", "2025.10.09（周四", "2025-10-09"]
            ]
          }
        ]
        """);

        var request = new DocxResolveRequest(
            "stability-report",
            stabilityDataPath,
            protocolPath,
            qualityStandardPath,
            reportPath,
            inspectionReportPath,
            samplingPlanPath);

        var result = Resolver.Resolve(plan, request);

        Assert.True(result.Operations.Count >= 54);
        Assert.Contains("0", result.ResolvedCommentIds);
        Assert.Contains("1", result.ResolvedCommentIds);
        Assert.Contains("2", result.ResolvedCommentIds);
        Assert.Contains("4", result.ResolvedCommentIds);
        Assert.Contains("8", result.ResolvedCommentIds);
        Assert.Contains("10", result.ResolvedCommentIds);
        Assert.Contains("14", result.ResolvedCommentIds);
        Assert.Contains("15", result.ResolvedCommentIds);
        Assert.Contains("16", result.ResolvedCommentIds);
        Assert.Contains("18", result.ResolvedCommentIds);
        Assert.Contains("20", result.ResolvedCommentIds);
        Assert.Contains("22", result.ResolvedCommentIds);
        Assert.Empty(result.UnresolvedItems);

        var intro = Assert.Single(result.Operations, op => op.Type == "replaceAnchoredText" && op.CommentId == "0");
        Assert.Equal("HSPXXXXDS工程批原液（250810S）", intro.Text);

        var narrative = Assert.Single(result.Operations, op => op.Type == "replaceParagraphText" && op.ParagraphIndex == 656);
        Assert.Contains("考察HSPXXXXDS（250810S）", narrative.Text);

        var photoIntro = Assert.Single(result.Operations, op => op.Type == "replaceParagraphText" && op.ParagraphIndex == 838);
        Assert.Contains("考察HSPXXXXDS（250810S）", photoIntro.Text);

        var summaryParagraph = Assert.Single(result.Operations, op => op.Type == "replaceParagraphText" && op.ParagraphIndex == 1069);
        Assert.Contains("HSPXXXXDS在高温", summaryParagraph.Text);

        Assert.Contains(result.Operations, op =>
            op.Type == "replaceTableCellText" &&
            op.TableIndex == 9 &&
            op.RowIndex == 2 &&
            op.CellIndex == 1 &&
            op.Text == "与黄色4号标准比色液比较，不得更深");

        Assert.Contains(result.Operations, op =>
            op.Type == "replaceTableCellText" &&
            op.TableIndex == 9 &&
            op.RowIndex == 4 &&
            op.CellIndex == 2 &&
            op.Text == "低于0.5号浊度标准液\nLower than reference suspension Ⅰ.");

        Assert.Contains(result.Operations, op =>
            op.Type == "replaceTableCellText" &&
            op.TableIndex == 9 &&
            op.RowIndex == 6 &&
            op.CellIndex == 2 &&
            op.Text == "5.6");

        Assert.Contains(result.Operations, op =>
            op.Type == "replaceTableCellText" &&
            op.TableIndex == 9 &&
            op.RowIndex == 7 &&
            op.CellIndex == 2 &&
            op.Text == "10.3 \nmg/ml");

        Assert.Contains(result.Operations, op =>
            op.Type == "replaceTableCellText" &&
            op.TableIndex == 4 &&
            op.RowIndex == 0 &&
            op.CellIndex == 1 &&
            op.Text == "HSPXXXXDS");

        Assert.Contains(result.Operations, op =>
            op.Type == "replaceTableCellText" &&
            op.TableIndex == 4 &&
            op.RowIndex == 1 &&
            op.CellIndex == 1 &&
            op.Text == "2025-09-08");

        Assert.Contains(result.Operations, op =>
            op.Type == "replaceTableCellText" &&
            op.TableIndex == 4 &&
            op.RowIndex == 5 &&
            op.CellIndex == 1 &&
            op.Text == "3 ml/瓶3 mL/bottle");

        Assert.Contains(result.Operations, op =>
            op.Type == "replaceTableCellText" &&
            op.TableIndex == 2 &&
            op.RowIndex == 2 &&
            op.CellIndex == 2 &&
            op.Text == "2025-09-23");

        Assert.Contains(result.Operations, op =>
            op.Type == "replaceTableCellText" &&
            op.TableIndex == 2 &&
            op.RowIndex == 5 &&
            op.CellIndex == 3 &&
            op.Text == "2025-10-02");

        Assert.Contains(result.Operations, op =>
            op.Type == "replaceTableCellText" &&
            op.TableIndex == 2 &&
            op.RowIndex == 10 &&
            op.CellIndex == 4 &&
            op.Text == "2025-10-09");
    }

    private static DocxPlanResult CreateAna03Plan()
        => new(
            Input: "/tmp/ana03-report.docx",
            Scenario: "stability-report",
            Items:
            [
                CreatePlanItem("0", "source_mapping", "paragraph", 105, null),
                CreatePlanItem("1", "source_mapping", "current_table", 117, 2),
                CreatePlanItem("2", "source_mapping", "current_table", 119, 2),
                CreatePlanItem("4", "source_mapping", "current_table", 121, 2),
                CreatePlanItem("8", "source_mapping", "current_table", 305, null),
                CreatePlanItem("10", "source_mapping", "current_table", 323, 4),
                CreatePlanItem("14", "fill_table_block", "current_table", 507, 9),
                CreatePlanItem("15", "fill_table_block", "current_table", 509, 9),
                CreatePlanItem("16", "fill_table_block", "current_table", 511, 9),
                CreatePlanItem("18", "generate_paragraph", "section", 656, null)
                ,CreatePlanItem("20", "generate_paragraph", "section", 837, null)
                ,CreatePlanItem("22", "generate_paragraph", "section", 1069, null)
            ],
            ProposedEdits: [],
            Warnings: [],
            Confidence: "medium");

    private static DocxPlanItem CreatePlanItem(string commentId, string instructionType, string targetScope, int paragraphIndex, int? tableIndex)
        => new(
            CommentId: commentId,
            CommentText: commentId,
            Anchor: new AnnotationAnchor(
                CommentId: commentId,
                Author: "tester",
                CommentText: commentId,
                AnchorText: commentId,
                Source: "mainDocument",
                TargetKind: tableIndex is null ? "paragraph" : "tableCell",
                ParagraphIndex: paragraphIndex,
                TableIndex: tableIndex,
                RowIndex: tableIndex is null ? null : 0,
                CellIndex: tableIndex is null ? null : 0,
                NearestHeadingText: tableIndex is null ? "试验结果Testing Results" : "试验结果Testing Results",
                CurrentParagraphText: commentId switch
                {
                    "18" => "考察HSPXXXXDS（YYYY）在25 ℃±2 ℃/60%RH±5%RH高温条件下放置1个月、2个月、3个月的稳定性情况，结果汇总如表9所示。SEC-HPLC、icIEF、nrCE-SDS和rCE-SDS检项的叠图如图5-图8所示。",
                    "22" => "HSPXXXXDS在高温（25 ℃±2 ℃/60%RH±5%RH）条件下的影响因素研究结果显示：颜色、澄清度、pH、蛋白质含量、rCE-SDS、结合活性Ⅰ和结合活性Ⅱ检测结果无明显变化，均符合质量标准。SEC-HPLC、icIEF和nrCE-SDS纯度发生了一定程度的降低，但仍在质量标准范围内。本影响因素试验结果提示HSPXXXXDS对高温轻微敏感，应避免高温条件长时间储存。",
                    _ => null
                },
                PreviousParagraphText: null,
                FollowingParagraphText: commentId switch
                {
                    "20" => "考察HSPXXXXDS（YYYY）在温度为25 ℃±2 ℃，100%日光+100%紫外（13天总照度不低于1.2×106 Lux·hr，近紫外总能量≥ 200 w∙hr/m2）的条件下放置6天和13天的稳定性情况，同时在温度为25 ℃±2 ℃条件下进行光照阴性对照试验，结果汇总如表10所示，其中光照和光照阴性对照条件下...",
                    _ => null
                },
                CurrentTableRowCount: tableIndex is null ? null : 20,
                CurrentTableColumnCount: tableIndex is null ? null : 5),
            InstructionType: instructionType,
            TargetScope: targetScope,
            CandidateTargets: [],
            RequiredSources: [],
            Confidence: "medium",
            Reasoning: "test",
            ProposedEdits: []);

    private static string WriteTempJson(string content)
    {
        var path = Path.Combine(Path.GetTempPath(), $"ana03-resolve-{Guid.NewGuid():N}.json");
        File.WriteAllText(path, content);
        return path;
    }
}
