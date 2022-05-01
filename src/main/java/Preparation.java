public class Preparation {

    public static void run() {
        PropertiesLoader propertiesLoader = new PropertiesLoader("setup.properties");
        Config.INPUT_FILE_NAME = propertiesLoader.loadProperty("inputFile");
        Config.OUTPUT_FILE_NAME = propertiesLoader.loadProperty("outputFile");
        Config.INPUT_SHEET_NAME = propertiesLoader.loadProperty("inputSheet");
        Config.OUTPUT_SHEET_NAME = propertiesLoader.loadProperty("outputSheet");
        Config.ROW_BEGIN = Integer.parseInt(propertiesLoader.loadProperty("rowBegin"));
        Config.ROW_END = Integer.parseInt(propertiesLoader.loadProperty("rowEnd"));
    }

    public static void main(String[] args) {
        run();
        System.out.println(Config.INPUT_FILE_NAME);
        System.out.println(Config.OUTPUT_FILE_NAME);
        System.out.println(Config.INPUT_SHEET_NAME);
        System.out.println(Config.OUTPUT_SHEET_NAME);
    }
}
