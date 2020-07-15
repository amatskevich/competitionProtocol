package by.matskevich.protocol;

import by.matskevich.protocol.model.InputParams;
import by.matskevich.protocol.service.CreationService;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

@RestController
public class ProtocolController {

    private final CreationService service = new CreationService();

    @PostMapping("/protocol")
    public ResponseEntity<byte[]> generateProtocol(@RequestBody InputParams params) throws IOException {
        System.out.println("Start program.");
        System.out.println(params);
        Workbook workbook = service.generateProtocol(params);

        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        try {
            workbook.write(bos);
        } finally {
            bos.close();
            workbook.close();
        }

        byte[] bytes = bos.toByteArray();
        System.out.println("Finish program.");
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment;filename=protocol.xlsx")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .contentLength(bytes.length)
                .body(bytes);
    }

    @GetMapping("/info")
    public String getInfo() {
        return "Hello world!";
    }
}
