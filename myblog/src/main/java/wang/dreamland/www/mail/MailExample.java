package wang.dreamland.www.mail;

/**
 * Created by wly on 2018/3/7.
 */

public class MailExample {

    public static void main (String args[]) throws Exception {
        String email = "rb530005265@163.com";
        String validateCode = "sd";
        SendEmail.sendEmailMessage(email,validateCode);

    }
}
