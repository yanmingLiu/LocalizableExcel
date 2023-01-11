plugins {
    id("java-library")
//    id("org.jetbrains.kotlin.jvm")
}

dependencies {
    implementation(fileTree("dir" to "libs", "include" to listOf("*.jar", "*.aar")))
    implementation("com.alibaba:easyexcel:3.1.1")

}

tasks.withType<Javadoc>().all {
    options.encoding = "UTF-8"
}

java {
    targetCompatibility = JavaVersion.VERSION_1_7
    sourceCompatibility = JavaVersion.VERSION_1_7
}