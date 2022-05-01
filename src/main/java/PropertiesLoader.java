import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.io.Reader;
import java.util.Properties;

public class PropertiesLoader {
    private Properties properties = new Properties();
    private Reader reader = null;

    public PropertiesLoader(String path) {
        try {
            reader = new FileReader(path);
            try {
                properties.load(reader);
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        } catch (FileNotFoundException ex) {
            ex.printStackTrace();
        } finally {
            if (reader != null) {
                try {
                    reader.close();
                } catch (IOException ex) {
                    ex.printStackTrace();
                }
            }
        }
    }

    public String loadProperty(String propertyName) {
        return properties.getProperty(propertyName);
    }

}
