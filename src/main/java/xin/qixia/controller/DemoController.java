package xin.qixia.controller;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.core.io.ResourceLoader;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import xin.qixia.domain.ExcelBo;
import xin.qixia.domain.SysOperLog;
import xin.qixia.utils.ExcelUtils;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Type;
import java.util.ArrayList;
import java.util.List;

/**
 * @author qixia
 */
@RestController
@RequestMapping
public class DemoController {

    private final ResourceLoader resourceLoader;

    public DemoController(ResourceLoader resourceLoader) {
        this.resourceLoader = resourceLoader;
    }

    @GetMapping("/")
    public String index() {
        return "Hello World";
    }

    @GetMapping("/download")
    public void download(HttpServletResponse response) {

        InputStream inputStream;
        try {
            List<ExcelBo> excelBoList = new ArrayList<>();

            inputStream = resourceLoader.getResource("classpath:json/sys_oper_log.json").getInputStream();

            List<SysOperLog> list = new ObjectMapper().readValue(inputStream,
                    new TypeReference<>() {
                        @Override
                        public Type getType() {
                            return super.getType();
                        }
                    });

            for (int i = 0; i < 2; i++) {
                ExcelBo excelBo = new ExcelBo();
                if (i == 0) {
                    excelBo.setSheetName("Sheet1")
                            .put("title", "第一个子表");
                } else {
                    excelBo.setSheetName("Sheet2")
                            .put("title", "第二个子表");

                    excelBo.add(3, 4, 2, 4);
                }

                excelBo.put("data1", list);
                excelBoList.add(excelBo);
            }


            ExcelUtils.exportTemplateComplexity(excelBoList, "excel/示例.xlsx", response.getOutputStream());
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
