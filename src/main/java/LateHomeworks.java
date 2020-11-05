import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.classroom.ClassroomScopes;
import com.google.api.services.classroom.model.*;
import com.google.api.services.classroom.Classroom;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.*;

public class LateHomeworks {
    private static final String APPLICATION_NAME = "Google Classroom API Java LateHomeworks";
    private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
    private static final String TOKENS_DIRECTORY_PATH = "tokens";

    /**
     * Global instance of the scopes required by this quickstart.
     * If modifying these scopes, delete your previously saved tokens/ folder.
     */
    private static final List<String> SCOPES = List.of(ClassroomScopes.CLASSROOM_COURSES_READONLY,
                                                       ClassroomScopes.CLASSROOM_COURSEWORK_STUDENTS_READONLY,
                                                       ClassroomScopes.CLASSROOM_ROSTERS_READONLY);
    private static final String CREDENTIALS_FILE_PATH = "/credentials.json";

    /**
     * Set of first and last names of students to be checked
     * Add comma separated student list in the parentheses, and remove the examples
     */
    private static Set<String> MyStudents = new HashSet<>(List.of("Example Name One", "Example Name Two"));

    private static HashMap<String, String> studentIDMap = new HashMap<>();


    private static HashMap<String, List<String>> lateMap = new HashMap<>();

    /**
     * Creates an authorized Credential object.
     * @param HTTP_TRANSPORT The network HTTP Transport.
     * @return An authorized Credential object.
     * @throws IOException If the credentials.json file cannot be found.
     */
    private static Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT) throws IOException {
        // Load client secrets.
        InputStream in = LateHomeworks.class.getResourceAsStream(CREDENTIALS_FILE_PATH);
        if (in == null) {
            throw new FileNotFoundException("Resource not found: " + CREDENTIALS_FILE_PATH);
        }
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

        // Build flow and trigger user authorization request.
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                .setDataStoreFactory(new FileDataStoreFactory(new java.io.File(TOKENS_DIRECTORY_PATH)))
                .setAccessType("offline")
                .build();
        LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build();
        return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
    }

    public static void main(String... args) throws IOException, GeneralSecurityException {
        // Build a new authorized API client service.
        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
        Classroom service = new Classroom.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
                .setApplicationName(APPLICATION_NAME)
                .build();

        // Create List of all courses
        List<Course> courses = getCourseList(service);
        if (courses == null || courses.size() == 0) {
            System.out.println("No courses found.");
            return;
        }
        for (Course course : courses) {

            // Create List of students enrolled in course
            List<Student> students = getStudentList(service, course);
            if (students == null || students.size() == 0) {
                continue;
            }

            // Create List of coursework for this course
            List<CourseWork> courseWorks = getCourseWorkList(service, course);
            if (courseWorks == null || courseWorks.size() == 0) {
                continue;
            }

            for (CourseWork courseWork : courseWorks) {
                List<StudentSubmission> submissions = getSubmissionList(service, course, courseWork);
                if (submissions == null || submissions.size() == 0) {
                    continue;
                }
                for (StudentSubmission submission : submissions) {
                    if(MyStudents.contains(studentIDMap.get(submission.getUserId()))){
                        if(isLate(submission) || completedByMostOfClass(submissions)){
                            try {
                                List<String> appendedList = lateMap.get(studentIDMap.get(submission.getUserId()));
                                try {
                                    appendedList.add(courseWork.getTitle());
                                }catch (UnsupportedOperationException e){
                                    e.printStackTrace();
                                }
                                lateMap.put(studentIDMap.get(submission.getUserId()), appendedList);
                            }catch (NullPointerException e){
                                List<String> lateList = new LinkedList<>();
                                lateList.add(courseWork.getTitle());
                                lateMap.put(studentIDMap.get(submission.getUserId()), lateList);
                            }
                        }
                    }
                }
            }
        }
        printToExcel();
    }

    private static void printToExcel(){
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Late Homeworks");

        int rowCount = 0;

        for (String student : lateMap.keySet()) {
            Row row = sheet.createRow(++rowCount);

            int columnCount = 0;
            Cell cell = row.createCell(columnCount);
            cell.setCellValue(student);

            for (String assignment : lateMap.get(student)) {
                cell = row.createCell(++columnCount);
                cell.setCellValue(assignment);
            }

        }


        try (FileOutputStream outputStream = new FileOutputStream("LateHomeworks.xlsx")) {
            workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static List<Course> getCourseList(Classroom service){
        ListCoursesResponse response = null;
        try {
            response = service.courses().list()
                    .setPageSize(10)
                    .execute();
        } catch (IOException e) {
            e.printStackTrace();
        }
        if(response == null){
            return null;
        }
        return response.getCourses();
    }

    private static List<Student> getStudentList(Classroom service, Course course){
        ListStudentsResponse studentResponse = null;
        try {
            studentResponse = service.courses().students().list(course.getId())
                    .setPageSize(10)
                    .execute();
        } catch (IOException e) {
            e.printStackTrace();
        }
        if(studentResponse == null){
            return null;
        }
        List<Student> students = studentResponse.getStudents();
        for(Student student : students){
            studentIDMap.put(student.getUserId(), student.getProfile().getName().getFullName());
        }
        return students;
    }

    private static List<CourseWork> getCourseWorkList(Classroom service, Course course){
        ListCourseWorkResponse courseWorkResponse;
        try {
            courseWorkResponse = service.courses().courseWork().list(course.getId())
                    .setPageSize(10)
                    .execute();
        } catch (IOException e) {
            return null;
            //e.printStackTrace();
        }
        if(courseWorkResponse == null){
            return null;
        }
        return courseWorkResponse.getCourseWork();
    }

    private static List<StudentSubmission> getSubmissionList(Classroom service, Course course, CourseWork courseWork){
        ListStudentSubmissionsResponse studentSubmissionsResponse = null;
        try {
            studentSubmissionsResponse = service.courses().courseWork().studentSubmissions().list(
                    course.getId(),
                    courseWork.getId()).execute();
        } catch (IOException e) {
            e.printStackTrace();
        }
        if(studentSubmissionsResponse == null){
            return null;
        }
        return studentSubmissionsResponse.getStudentSubmissions();
    }

    private static boolean isLate(StudentSubmission submission){
        if(submission.getLate() == null){
            return false;
        }
        return submission.getLate();
    }

    private static boolean completedByMostOfClass(List<StudentSubmission> submissions){
        int completedSubmissions = 0;
        for(StudentSubmission submission : submissions){
            if(submission.getState().equals("TURNED_IN")){
                completedSubmissions++;
            }
        }
        return completedSubmissions > 9;
    }
}
