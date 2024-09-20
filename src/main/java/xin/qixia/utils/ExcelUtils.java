package xin.qixia.utils;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.alibaba.excel.write.metadata.fill.FillWrapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import xin.qixia.domain.ExcelBo;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Collection;
import java.util.Collections;
import java.util.List;
import java.util.Map;

/**
 * @author qixia
 */
public class ExcelUtils {
    /**
     * 复杂数据模板导出
     *
     * @param excelBoList 模板需要的数据
     */
    private static void exportTemplateComplexity(List<ExcelBo> excelBoList, ExcelWriter excelWriter) {
        if (excelBoList == null || excelBoList.isEmpty()) {
            throw new IllegalArgumentException("数据为空");
        }
        for (ExcelBo excelBo : excelBoList) {
            WriteSheet writeSheet = EasyExcel.writerSheet(excelBo.getSheetName()).build();
            for (Map.Entry<String, Object> map : excelBo.getData().entrySet()) {
                // 设置列表后续还有数据
                FillConfig fillConfig = FillConfig.builder().forceNewRow(Boolean.TRUE).build();
                if (map.getValue() instanceof Collection) {
                    // 多表导出必须使用 FillWrapper
                    excelWriter.fill(new FillWrapper(map.getKey(), (Collection<?>) map.getValue()), fillConfig, writeSheet);
                } else {
                    excelWriter.fill(map.getValue(), writeSheet);
                }
            }
            //合并单元格
            if (excelBo.getCellList() != null && !excelBo.getCellList().isEmpty()) {
                Sheet sheet = excelWriter.writeContext().writeSheetHolder().getSheet();
                for (CellRangeAddress cellAddresses : excelBo.getCellList()) {
                    sheet.addMergedRegion(cellAddresses);
                }
            }
        }
        excelWriter.finish();
    }

    /**
     * 复杂数据模板导出
     *
     * @param excelBoList  模板需要的数据
     * @param templatePath 模板路径 resource 目录下的路径包括模板文件名
     *                     例如: excel/temp.xlsx
     *                     重点: 模板文件必须放置到启动类对应的 resource 目录下
     * @param os           输出流
     */
    public static void exportTemplateComplexity(List<ExcelBo> excelBoList, String templatePath, OutputStream os) {
        ClassPathResource templateResource = new ClassPathResource(templatePath);
        ExcelWriter excelWriter;
        try {
            excelWriter = EasyExcel.write(os)
                    .withTemplate(templateResource.getInputStream())
                    .autoCloseStream(false)
                    .build();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        exportTemplateComplexity(excelBoList, excelWriter);
    }

    /**
     * 复杂数据模板导出
     *
     * @param excelBo      模板需要的数据
     * @param templatePath 模板路径 resource 目录下的路径包括模板文件名
     *                     例如: excel/temp.xlsx
     *                     重点: 模板文件必须放置到启动类对应的 resource 目录下
     * @param os           输出流
     */
    public static void exportTemplateComplexity(ExcelBo excelBo, String templatePath, OutputStream os) {
        exportTemplateComplexity(Collections.singletonList(excelBo), templatePath, os);
    }

    /**
     * 复杂数据模板导出
     *
     * @param excelBoList  模板需要的数据
     * @param templateFile 模板文件
     * @param os           输出流
     */
    public static void exportTemplateComplexity(List<ExcelBo> excelBoList, File templateFile, OutputStream os) {
        ExcelWriter excelWriter = EasyExcel.write(os)
                .withTemplate(templateFile)
                .autoCloseStream(false)
                .build();
        exportTemplateComplexity(excelBoList, excelWriter);
    }

    /**
     * 复杂数据模板导出
     *
     * @param excelBo      模板需要的数据
     * @param templateFile 模板文件
     * @param os           输出流
     */
    public static void exportTemplateComplexity(ExcelBo excelBo, File templateFile, OutputStream os) {
        exportTemplateComplexity(Collections.singletonList(excelBo), templateFile, os);
    }

    /**
     * 获取模板内容和格式
     *
     * @param sourceSheet
     * @param targetSheet
     */
    private static void copySheet(Sheet sourceSheet, Sheet targetSheet) {
        Workbook targetWorkbook = targetSheet.getWorkbook();
        CellStyle targetCellStyle;

        // 复制合并的单元格区域
        for (int i = 0; i < sourceSheet.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = sourceSheet.getMergedRegion(i);
            targetSheet.addMergedRegion(mergedRegion);
        }

        for (Row sourceRow : sourceSheet) {
            Row targetRow = targetSheet.createRow(sourceRow.getRowNum());

            for (Cell sourceCell : sourceRow) {
                Cell targetCell = targetRow.createCell(sourceCell.getColumnIndex());

                // 复制单元格的值
                switch (sourceCell.getCellType()) {
                    case STRING:
                        targetCell.setCellValue(sourceCell.getStringCellValue());
                        break;
                    case NUMERIC:
                        targetCell.setCellValue(sourceCell.getNumericCellValue());
                        break;
                    case BOOLEAN:
                        targetCell.setCellValue(sourceCell.getBooleanCellValue());
                        break;
                }

                // 复制单元格的样式
                CellStyle sourceCellStyle = sourceCell.getCellStyle();
                targetCellStyle = targetWorkbook.createCellStyle();
                targetCellStyle.cloneStyleFrom(sourceCellStyle);
                targetCell.setCellStyle(targetCellStyle);
            }
        }
    }

    /**
     * 生成多sheet模板
     *
     * @param outModelFile
     * @param templateFile
     */
    private static void createModel(File outModelFile, String templateFile, List<ExcelBo> excelBoList) {
        try (FileOutputStream fileOutputStream = new FileOutputStream(outModelFile); Workbook targetWorkbook = new XSSFWorkbook()) {
            ClassPathResource templateResource = new ClassPathResource(templateFile);
            //读取excel模板
            XSSFWorkbook workbook = new XSSFWorkbook(templateResource.getInputStream());
            workbook.cloneSheet(0);

            // 获取要复制的工作表索引
            int sourceSheetIndex = 0;
            Sheet sourceSheet = workbook.getSheetAt(sourceSheetIndex);

            for (ExcelBo excelBo : excelBoList) {
                Sheet targetSheet = targetWorkbook.createSheet(excelBo.getSheetName());
                // 复制原始工作表的内容到目标工作表
                copySheet(sourceSheet, targetSheet);
            }
            // 保存目标文件
            targetWorkbook.write(fileOutputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 单Sheet模板导出到多个Sheet 模板格式为 {key.属性}
     *
     * @param excelBoList  模板需要的数据
     * @param templatePath 模板路径 resource 目录下的路径包括模板文件名
     *                     例如: excel/temp.xlsx
     *                     重点: 模板文件必须放置到启动类对应的 resource 目录下
     * @param os           输出流
     */
    public static void exportTemplateMultiSheet(List<ExcelBo> excelBoList, String templatePath, OutputStream os) {
        File outModelFile = new File(System.currentTimeMillis() + ".xlsx");
        //创建模板
        createModel(outModelFile, templatePath, excelBoList);
        //复杂数据模板导出
        exportTemplateComplexity(excelBoList, outModelFile, os);
        //删除模板
        if (outModelFile.exists()) {
            if (outModelFile.delete()) {
                throw new RuntimeException("file delete succeed [" + outModelFile.getPath() + "]");
            } else {
                throw new RuntimeException("file delete fail [" + outModelFile.getPath() + "]");
            }
        }
    }
}
