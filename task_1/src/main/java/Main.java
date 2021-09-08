import java.io.IOException;

public class Main {
    public static void main(String[] args) throws IOException {
        try {
            CreateExcel.CreateSimple(args);
        } catch (Throwable throwable) {
            throwable.printStackTrace();
        }
    }
}
