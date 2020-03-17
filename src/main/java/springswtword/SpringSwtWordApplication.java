package springswtword;

import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class SpringSwtWordApplication implements CommandLineRunner {

    public static void main(String[] args) {
        SpringApplication.run(SpringSwtWordApplication.class, args);
    }

    @Override
    public void run(String... args) throws Exception {
        new WordFrame().openning();
    }
}
