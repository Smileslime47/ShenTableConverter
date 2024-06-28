plugins {
    kotlin("jvm") version "1.8.20"
    application
}

group = "moe._47saikyo"
version = "1.0-SNAPSHOT"

repositories {
    maven ("https://maven.aliyun.com/repository/google")
    maven ("https://maven.aliyun.com/repository/public/")
    mavenCentral()
    google()
}

dependencies {
    implementation("org.apache.poi:poi:5.2.5")
    implementation("org.apache.poi:poi-ooxml:5.2.5")
    testImplementation(kotlin("test"))
}

tasks.test {
    useJUnitPlatform()
}

kotlin {
    jvmToolchain(11)
}

application {
    mainClass.set("MainKt")
}