package xin.qixia;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;
import xin.qixia.domain.SysOperLog;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Type;
import java.util.List;

@SpringBootTest
class DemoApplicationTests {

    @Test
    void contextLoads() throws IOException {
        ObjectMapper objectMapper = new ObjectMapper();
        List<SysOperLog> sysOperLog = objectMapper.readValue(
                new File("D:\\IdeaProjects\\tool\\excel-tool\\src\\main\\resources\\json\\sys_oper_log.json"),
                new TypeReference<>() {
                    @Override
                    public Type getType() {
                        return super.getType();
                    }
                });
        sysOperLog.forEach(System.out::println);
    }

}
