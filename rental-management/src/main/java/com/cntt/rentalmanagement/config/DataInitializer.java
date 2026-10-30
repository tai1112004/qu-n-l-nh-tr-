package com.cntt.rentalmanagement.config;

import com.cntt.rentalmanagement.domain.enums.AuthProvider;
import com.cntt.rentalmanagement.domain.enums.RoleName;
import com.cntt.rentalmanagement.domain.models.Category;
import com.cntt.rentalmanagement.domain.models.Location;
import com.cntt.rentalmanagement.domain.models.Role;
import com.cntt.rentalmanagement.domain.models.User;
import com.cntt.rentalmanagement.repository.CategoryRepository;
import com.cntt.rentalmanagement.repository.LocationRepository;
import com.cntt.rentalmanagement.repository.RoleRepository;
import com.cntt.rentalmanagement.repository.UserRepository;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.boot.CommandLineRunner;
import org.springframework.security.crypto.password.PasswordEncoder;
import org.springframework.stereotype.Component;

import java.util.HashSet;
import java.util.Optional;
import java.util.Set;

@Component
@RequiredArgsConstructor
@Slf4j
public class DataInitializer implements CommandLineRunner {

    private final UserRepository userRepository;
    private final RoleRepository roleRepository;
    private final PasswordEncoder passwordEncoder;
    private final CategoryRepository categoryRepository;
    private final LocationRepository locationRepository;

    @Override
    public void run(String... args) {
        try {
            // 1. Initialize Roles if not present
            Role adminRole = getOrCreateRole(RoleName.ROLE_ADMIN);
            getOrCreateRole(RoleName.ROLE_RENTALER);
            getOrCreateRole(RoleName.ROLE_USER);

            // 2. Initialize Default Admin Account (admin@gmail.com / admin@123)
            initAdminUser(adminRole);

            // 3. Initialize default location & category if empty
            initLocationAndCategory();

        } catch (Exception ex) {
            log.error("Error during initial data setup: {}", ex.getMessage(), ex);
        }
    }

    private Role getOrCreateRole(RoleName roleName) {
        return roleRepository.findByName(roleName).orElseGet(() -> {
            Role role = new Role();
            role.setName(roleName);
            Role saved = roleRepository.save(role);
            log.info("Initialized role: {}", roleName);
            return saved;
        });
    }

    private void initAdminUser(Role adminRole) {
        String adminEmail = "admin@gmail.com";
        Optional<User> userOpt = userRepository.findByEmail(adminEmail);
        
        if (userOpt.isPresent()) {
            User existingAdmin = userOpt.get();
            existingAdmin.setPassword(passwordEncoder.encode("admin@123"));
            existingAdmin.setIsLocked(false);
            existingAdmin.setIsConfirmed(true);
            existingAdmin.setEmailVerified(true);
            
            Set<Role> roles = existingAdmin.getRoles();
            if (roles == null) {
                roles = new HashSet<>();
            }
            roles.add(adminRole);
            existingAdmin.setRoles(roles);
            
            userRepository.save(existingAdmin);
            log.info("Default Admin account updated: email='{}', password='admin@123'", adminEmail);
        } else {
            User newAdmin = new User();
            newAdmin.setName("Administrator");
            newAdmin.setEmail(adminEmail);
            newAdmin.setPassword(passwordEncoder.encode("admin@123"));
            newAdmin.setPhone("0999999999");
            newAdmin.setIsLocked(false);
            newAdmin.setIsConfirmed(true);
            newAdmin.setEmailVerified(true);
            newAdmin.setProvider(AuthProvider.local);

            Set<Role> roles = new HashSet<>();
            roles.add(adminRole);
            newAdmin.setRoles(roles);

            userRepository.save(newAdmin);
            log.info("Default Admin account created successfully: email='{}', password='admin@123'", adminEmail);
        }
    }

    private void initLocationAndCategory() {
        if (locationRepository.count() == 0) {
            Location loc = new Location();
            loc.setCityName("Hà Nội");
            locationRepository.save(loc);
            log.info("Initialized default location: Hà Nội");
        }

        if (categoryRepository.count() == 0) {
            createCategory("Bất động sản");
            createCategory("Phòng trọ");
            createCategory("Chung cư mini");
            log.info("Initialized default categories.");
        }
    }

    private void createCategory(String name) {
        Category cat = new Category();
        cat.setName(name);
        categoryRepository.save(cat);
    }
}
