import com.azure.identity.*;
import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonObject;
import com.microsoft.graph.authentication.TokenCredentialAuthProvider;
import com.microsoft.graph.logger.DefaultLogger;
import com.microsoft.graph.models.*;
import com.microsoft.graph.requests.GraphServiceClient;
import com.microsoft.graph.serializer.DefaultSerializer;
import com.mysql.cj.jdbc.result.ResultSetImpl;
import net.minidev.json.JSONObject;
import org.apache.http.*;
import org.apache.http.client.entity.UrlEncodedFormEntity;
import org.apache.http.client.methods.HttpPatch;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.ContentType;
import org.apache.http.entity.StringEntity;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.message.BasicNameValuePair;
import org.apache.http.util.EntityUtils;

import java.io.*;
import java.sql.*;
import java.util.*;
import java.util.List;


public class Main {
    private final static CloseableHttpClient httpClient = HttpClients.createDefault();
    private static GraphServiceClient graphClient;
    private static String access_token;

    public static void main(String[] args) {

        final Properties propertiesPrivate = new Properties();
        final Properties propertiesNenomicsPersonal = new Properties();
        final Properties propertiesNenomics = new Properties();
        try {
            propertiesPrivate.load(Main.class.getResourceAsStream("private.properties"));
            propertiesNenomicsPersonal.load(Main.class.getResourceAsStream("neopersonal.properties"));
            propertiesNenomics.load(Main.class.getResourceAsStream("neonomics.properties"));
        } catch (IOException e) {
            System.out.println("Unable to read OAuth configuration. Make sure you have a properly " +
                    "formatted .properties file. See README for details.");
            return;
        }

        ResultSet rs;
        String sql = "SELECT ip, clientId, endpoint, date, paymentAmount, currency, category FROM billing limit 50";
        String rangeAddress = null;
        JsonArray values = null;
        try {
            Class.forName("com.mysql.cj.jdbc.Driver");
            Connection con = DriverManager.getConnection(
                    "jdbc:mysql://localhost:3309/billing",
                    "mferoce", propertiesNenomics.getProperty("billingPwd"));
            PreparedStatement p = con.prepareStatement(sql);
            rs = p.executeQuery();
            int nColumns = rs.getMetaData().getColumnCount(); 
            int nRows = ((ResultSetImpl) rs).getRows().size();
            rangeAddress = getRangeAddressFromRs("CK1", rs, nColumns, nRows);
            values = buildJsonArrayValues(rs, nColumns);
            con.close();
        } catch (ClassNotFoundException | SQLException e) {
            e.printStackTrace();
        }

        JsonArray formulas = null;
        JsonArray formats = null;
        Map<String,JsonArray> arrays = new HashMap<>();
        arrays.put("values", values);
        //Can also add formulas, formats
        JsonObject patchBody = JsonObjectFromArrays(arrays);
        //String rangeAddress = "CK1:CO4";

        //sdkAuthUpdate(propertiesNenomics, rangeAddress, patchBody, values, formulas, formats);
        authenticateAndUpdate(propertiesNenomics, rangeAddress, patchBody);
    }

    private static String getRangeAddressFromRs(String firstCell, ResultSet rs, int nColumns, int nRows) throws SQLException {
        String[] columnsRows = firstCell.split("(?<=\\D)(?=\\d)|(?<=\\d)(?=\\D)");
        String newRow = String.valueOf(Integer.parseInt(columnsRows[1]) + nRows-1);
        String newColumn = addIntToColumnValue(columnsRows[0], nColumns-1);
        return firstCell+":"+newColumn+newRow;
    }

    private static String addIntToColumnValue(String s, int toAdd) {
        char[] charArr = s.toCharArray();
        char newLetSec;
        char newLetFirst;
        int newLetIntSec;
        int newLetIntFirst;
        int toAddFirstLetter=0;
        //A = 10
        //Z = 35

        newLetIntSec = Character.getNumericValue(charArr[charArr.length -1])+toAdd;
        while(newLetIntSec > 35){
            newLetIntSec -= 25;
            toAddFirstLetter++;
        }
        newLetIntFirst = Character.getNumericValue(charArr[0])+toAddFirstLetter;
        newLetSec = Character.toUpperCase(Character.forDigit(newLetIntSec, 36));
        newLetFirst = Character.toUpperCase(Character.forDigit(newLetIntFirst, 36));

        return newLetFirst +Character.toString(newLetSec);
    }

    private static JsonArray buildJsonArrayValues(ResultSet rs, int columnSize) throws SQLException {
        JsonArray values = new JsonArray();
        JsonArray jsonArr;
        
        while (rs.next()){
            jsonArr = new JsonArray(columnSize);
            for (int i = 1; i <= columnSize; i++) {
                jsonArr.add(rs.getString(i));
            }
            values.add(jsonArr);
        }
        return values;
    }

    private static void sdkAuthUpdate(Properties properties, String rangeAddress, JsonObject patchBody,
                                      JsonArray values, JsonArray formulas, JsonArray formats) {
        if(graphClient==null) graphClient = getGraphClient(properties);

        updateRangeExcelSdk(patchBody,values, formulas, formats, rangeAddress, graphClient, properties);
    }

    private static void updateRangeExcelSdk(JsonObject patchBody, JsonArray values, JsonArray formulas, JsonArray formats,
                                            String rangeAddress, GraphServiceClient graphClient,
                                            Properties properties) {
        WorkbookWorksheetRangeParameterSet rangeSet = new WorkbookWorksheetRangeParameterSet();
        rangeSet.address = rangeAddress;
        WorkbookRangeFill workbookRangeFill = new WorkbookRangeFill();
        workbookRangeFill.setRawObject(new DefaultSerializer(new DefaultLogger()),patchBody);
        graphClient.drives("b!OwIdmgGyD0GFLmOh-pY2pTyNwpj56XtHijy3Fdp_QriBJuhLBnh1Q5NrH1b1svpi")
                .items("01IQ5WS73QW5W2QOQN3RC33FWMPREHI766").workbook().worksheets("RawData").range(rangeSet)
                        .format().fill().buildRequest().patch(workbookRangeFill);

    }

    private static GraphServiceClient getGraphClient(Properties properties) {
        ClientSecretCredential clientSecretCredential = new ClientSecretCredentialBuilder()
                .clientId(properties.getProperty("app.clientId"))
                .clientSecret(properties.getProperty("app.clientSecret"))
                .tenantId(properties.getProperty("app.tenantId"))
                .build();


        List<String> scopes = new ArrayList<>();
        scopes.add("https://graph.microsoft.com/.default");
        TokenCredentialAuthProvider tokenCredentialAuthProvider = new TokenCredentialAuthProvider(scopes, clientSecretCredential);

        GraphServiceClient graphClient =
                GraphServiceClient
                        .builder()
                        .authenticationProvider(tokenCredentialAuthProvider)
                        .buildClient();
        return graphClient;
    }

    private static void authenticateAndUpdate(Properties properties, String rangeAddress, JsonObject patchBody) {
        if(access_token==null) {
            try {
                access_token = getAccessToken(properties);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        HttpResponse response = null;
        try {
            response = updateRangeExcel(properties, access_token, rangeAddress, patchBody);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static JsonObject JsonObjectFromArrays(Map<String,JsonArray> arrays) {
        JsonObject gson = new Gson().fromJson("{}", JsonObject.class);
        for (Map.Entry<String, JsonArray> arr: arrays.entrySet()) {
            gson.add(arr.getKey(),arr.getValue());
        }
        return gson;
    }

    private static HttpResponse updateRangeExcel(Properties properties, String token, String rangeAddress,
                                                 JsonObject patchBody) throws Exception {
        if(patchBody.get("values") == null) throw new Exception("values can't be null");
        HttpPatch httpPatch = new HttpPatch("https://graph.microsoft.com/v1.0/drives/"
                +properties.getProperty("dokumenterDriveId")+"/items/"+properties.getProperty("dailyTrafficReportId")
                +"/workbook/worksheets/" + properties.getProperty("app.sheetName")+"/range(address='"+rangeAddress+"')");
        httpPatch.addHeader("Content-Type", "application/json");
        httpPatch.addHeader("Authorization", "bearer "+token);
        httpPatch.setEntity(new StringEntity(patchBody.toString(), ContentType.APPLICATION_JSON));
        return httpClient.execute(httpPatch);
    }

    private static String getAccessToken(Properties properties) throws IOException {
        //https://docs.microsoft.com/en-us/graph/auth-v2-service
        HttpPost httppost = new HttpPost("https://login.microsoftonline.com/"
                +properties.getProperty("app.tenantId")+"/oauth2/v2.0/token");


        List<NameValuePair> form = new ArrayList<>() {{
            add(new BasicNameValuePair("grant_type", "client_credentials"));
            add(new BasicNameValuePair("client_id", properties.getProperty("app.clientId")));
            add(new BasicNameValuePair("client_secret", properties.getProperty("app.clientSecret")));
            add(new BasicNameValuePair("scope", "https://graph.microsoft.com/.default"));
        }};

        UrlEncodedFormEntity entity = new UrlEncodedFormEntity(form, Consts.UTF_8);

        httppost.setEntity(entity);
        httppost.addHeader("Content-Type", "application/x-www-form-urlencoded");

        return (String) new Gson().fromJson(EntityUtils.toString(httpClient.execute(httppost).getEntity()),
                JSONObject.class).get("access_token");
    }
}
