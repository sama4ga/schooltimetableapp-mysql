-- MySQL dump 10.13  Distrib 8.0.11, for Win64 (x86_64)
--
-- Host: localhost    Database: timetable
-- ------------------------------------------------------
-- Server version	8.0.11

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
 SET NAMES utf8 ;
/*!40103 SET @OLD_TIME_ZONE=@@TIME_ZONE */;
/*!40103 SET TIME_ZONE='+00:00' */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

--
-- Table structure for table `class`
--

DROP TABLE IF EXISTS `class`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
 SET character_set_client = utf8mb4 ;
CREATE TABLE `class` (
  `class_id` int(11) NOT NULL AUTO_INCREMENT,
  `name` varchar(40) DEFAULT NULL,
  `subject_combination` enum('all','selected') DEFAULT 'all',
  PRIMARY KEY (`class_id`)
) ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci COMMENT='|Holds class details';
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `class`
--

LOCK TABLES `class` WRITE;
/*!40000 ALTER TABLE `class` DISABLE KEYS */;
INSERT INTO `class` VALUES (1,'Junior Cambridge 1','all'),(2,'Junior Cambridge 2','all'),(3,'Junior Cambridge 3','all'),(4,'Senior Cambridge 1','all'),(5,'Senior Cambridge 2','selected'),(6,'Senior Cambridge 3','selected');
/*!40000 ALTER TABLE `class` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `misc`
--

DROP TABLE IF EXISTS `misc`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
 SET character_set_client = utf8mb4 ;
CREATE TABLE `misc` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `details` varchar(45) DEFAULT NULL,
  `start_time` varchar(5) DEFAULT NULL,
  `end_time` varchar(5) DEFAULT NULL,
  `days` varchar(15) DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=21 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci COMMENT='Used to hold miscellaneous details';
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `misc`
--

LOCK TABLES `misc` WRITE;
/*!40000 ALTER TABLE `misc` DISABLE KEYS */;
INSERT INTO `misc` VALUES (1,'LUNCH BREAK','11:30','12:00','Monday'),(2,'LUNCH BREAK','11:30','12:00','Tuesday'),(3,'LUNCH BREAK ','11:30','12:00','Wednesday'),(4,'LUNCH BREAK','11:30','12:00','Thursday'),(5,'LUNCH BREAK','11:30','12:00','Friday'),(8,'B-CLUB','10:50','11:30','Monday'),(9,'FITNESS','8:10','8:50','Friday'),(10,'FITNESS','8:50','9:30','Friday'),(11,'TENDER CHESS','14:00','14:30','Monday'),(12,'UCMAS','14:00','14:30','Tuesday'),(13,'MORALITY ETHICS','14:00','14:30','Wednesday'),(14,'CODING CLUB','14:00','14:30','Thursday'),(15,'MUSIC CLUB','14:00','14:30','Friday'),(16,'DEVOTION','7:45','8:10','Monday'),(17,'DEVOTION','7:45','8:10','Tuesday'),(18,'DEVOTION','7:45','8:10','Wednesday'),(19,'DEVOTION','7:45','8:10','Thursday'),(20,'DEVOTION','7:45','8:10','Friday');
/*!40000 ALTER TABLE `misc` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `subject_class`
--

DROP TABLE IF EXISTS `subject_class`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
 SET character_set_client = utf8mb4 ;
CREATE TABLE `subject_class` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `subject_id` int(11) DEFAULT NULL,
  `class_id` int(11) DEFAULT NULL,
  `time_per_period` int(3) DEFAULT NULL,
  `no_times_per_week` int(1) DEFAULT NULL,
  `no_double_period` int(1) DEFAULT NULL,
  `total_time_per_week` int(3) DEFAULT NULL,
  `category` varchar(10) DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=125 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci COMMENT='Holds subjects and the classes the subjects belongs to';
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `subject_class`
--

LOCK TABLES `subject_class` WRITE;
/*!40000 ALTER TABLE `subject_class` DISABLE KEYS */;
INSERT INTO `subject_class` VALUES (13,1,1,40,4,1,160,'General'),(14,1,2,40,4,1,160,'General'),(15,1,3,40,4,1,160,'General'),(16,1,4,40,3,1,120,'General'),(17,1,5,40,3,1,120,'General'),(18,1,6,40,3,1,120,'General'),(19,2,1,40,3,0,120,'General'),(20,2,2,40,3,0,120,'General'),(21,2,3,40,3,0,120,'General'),(22,2,4,40,3,0,120,'General'),(23,2,5,40,3,1,120,'General'),(24,2,6,40,3,1,120,'General'),(25,3,1,40,2,0,80,'General'),(26,3,2,40,2,0,80,'General'),(27,3,3,40,2,0,80,'General'),(28,4,1,40,2,0,80,'General'),(29,4,2,40,2,0,80,'General'),(30,4,3,40,2,0,80,'General'),(31,5,1,40,2,0,80,'General'),(32,5,2,40,2,0,80,'General'),(33,5,3,40,2,0,80,'General'),(34,6,1,40,2,0,80,'General'),(35,6,2,40,2,0,80,'General'),(36,6,3,40,2,0,80,'General'),(43,9,1,40,1,0,40,'General'),(44,9,2,40,1,0,40,'General'),(45,9,3,40,1,0,40,'General'),(46,10,1,40,1,0,40,'General'),(47,10,2,40,1,0,40,'General'),(48,10,3,40,1,0,40,'General'),(51,12,1,40,2,0,80,'General'),(52,12,2,40,2,0,80,'General'),(53,12,3,40,2,0,80,'General'),(54,13,1,40,2,0,80,'General'),(55,13,2,40,2,0,80,'General'),(56,13,3,40,2,0,80,'General'),(59,15,1,40,2,0,80,'General'),(60,15,2,40,2,0,80,'General'),(61,16,4,40,2,0,80,'General'),(62,16,5,40,3,1,120,'General'),(63,16,6,40,3,1,120,'General'),(64,17,4,40,2,0,80,'General'),(65,17,5,40,3,1,120,'Science'),(66,17,6,40,3,1,120,'Science'),(67,18,4,40,2,0,80,'General'),(68,18,5,40,3,1,120,'General'),(69,18,6,40,3,1,120,'General'),(70,19,4,40,2,0,80,'General'),(71,19,5,40,3,1,120,'General'),(72,19,6,40,3,1,120,'General'),(73,20,4,40,2,0,80,'General'),(74,20,5,40,2,0,80,'General'),(75,20,6,40,2,0,80,'General'),(76,21,1,40,2,0,80,'General'),(77,21,2,40,2,0,80,'General'),(78,21,4,40,2,0,80,'General'),(79,21,5,40,2,0,80,'Science'),(80,21,6,40,2,0,80,'Science'),(81,22,4,40,1,0,40,'General'),(82,22,5,40,1,0,40,'General'),(83,22,6,40,2,0,80,'General'),(84,23,4,40,2,0,80,'General'),(85,23,5,40,3,1,120,'Science'),(86,23,6,40,3,1,120,'Science'),(87,24,4,40,2,0,80,'General'),(88,24,5,40,2,0,80,'Art'),(89,24,6,40,2,0,80,'Art'),(90,25,4,40,2,0,80,'General'),(91,25,5,40,2,0,80,'Art'),(92,25,6,40,2,0,80,'Art'),(93,26,1,40,2,0,80,'Art'),(94,26,2,40,2,0,80,'Art'),(95,26,4,40,2,0,80,'Art'),(96,26,5,40,2,0,80,'Art'),(97,26,6,40,2,0,80,'Art'),(98,27,1,40,1,0,40,'General'),(99,27,2,40,1,0,40,'General'),(100,27,3,40,1,0,40,'General'),(101,27,4,40,1,0,40,'General'),(102,27,5,40,1,0,40,'General'),(103,27,6,40,1,0,40,'General'),(104,28,1,40,1,0,40,'Art'),(105,28,2,40,1,0,40,'Art'),(106,28,3,40,1,0,40,'Art'),(107,28,4,40,1,0,40,'Art'),(108,28,5,40,2,0,80,'Art'),(109,28,6,40,2,0,80,'Art'),(110,29,4,40,2,0,80,'General'),(111,29,5,40,2,0,80,'Commercial'),(112,29,6,40,2,0,80,'Commercial'),(113,30,4,40,2,0,80,'Commercial'),(114,30,5,40,2,0,80,'Commercial'),(115,30,6,40,2,0,80,'Commercial'),(116,31,4,40,2,0,80,'General'),(117,31,5,40,2,0,80,'Art'),(118,31,6,40,2,0,80,'Art'),(119,32,1,40,1,0,40,'General'),(120,32,2,40,1,0,40,'General'),(121,32,3,40,1,0,40,'General'),(122,33,4,40,2,0,80,'General'),(123,33,5,40,2,0,80,'General'),(124,33,6,40,2,0,80,'General');
/*!40000 ALTER TABLE `subject_class` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `subjects`
--

DROP TABLE IF EXISTS `subjects`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
 SET character_set_client = utf8mb4 ;
CREATE TABLE `subjects` (
  `subject_id` int(11) NOT NULL AUTO_INCREMENT,
  `name` varchar(100) DEFAULT NULL,
  PRIMARY KEY (`subject_id`)
) ENGINE=InnoDB AUTO_INCREMENT=34 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci COMMENT='Holds subject details';
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `subjects`
--

LOCK TABLES `subjects` WRITE;
/*!40000 ALTER TABLE `subjects` DISABLE KEYS */;
INSERT INTO `subjects` VALUES (1,'MATHS'),(2,'ENG. LANG.'),(3,'BST'),(4,'PVS'),(5,'RNV'),(6,'CCA'),(9,'MUSIC'),(10,'FRENCH'),(12,'IBIBIO'),(13,'BUSINESS STD.'),(15,'CAMBRIDGE SC.'),(16,'PHYSICS'),(17,'CHEMISTRY'),(18,'BIOLOGY'),(19,'ECONOMICS'),(20,'AGRIC. SC.'),(21,'GEOGRAPHY'),(22,'T. D.'),(23,'F. MATH'),(24,'CIVIC EDU.'),(25,'GOVT.'),(26,'HISTORY'),(27,'DICTION'),(28,'LITERATURE'),(29,'ACCOUNTING'),(30,'MARKETING'),(31,'CRS'),(32,'RNV-CRS'),(33,'ICT');
/*!40000 ALTER TABLE `subjects` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `teacher_available_time`
--

DROP TABLE IF EXISTS `teacher_available_time`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
 SET character_set_client = utf8mb4 ;
CREATE TABLE `teacher_available_time` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `teacher_id` int(11) DEFAULT NULL,
  `from` int(2) DEFAULT NULL,
  `to` int(2) DEFAULT NULL,
  `day` varchar(10) DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci COMMENT='Holds times teachers are available';
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `teacher_available_time`
--

LOCK TABLES `teacher_available_time` WRITE;
/*!40000 ALTER TABLE `teacher_available_time` DISABLE KEYS */;
INSERT INTO `teacher_available_time` VALUES (1,10,8,14,'Thursday'),(2,6,8,14,'Thursday');
/*!40000 ALTER TABLE `teacher_available_time` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `teacher_class`
--

DROP TABLE IF EXISTS `teacher_class`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
 SET character_set_client = utf8mb4 ;
CREATE TABLE `teacher_class` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `teacher_id` int(11) DEFAULT NULL,
  `class_id` int(11) DEFAULT NULL,
  `subject_id` int(11) DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=101 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci COMMENT='Holds details of classes handled by teachers';
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `teacher_class`
--

LOCK TABLES `teacher_class` WRITE;
/*!40000 ALTER TABLE `teacher_class` DISABLE KEYS */;
INSERT INTO `teacher_class` VALUES (1,1,1,1),(2,1,2,1),(3,1,3,1),(4,1,4,1),(5,1,5,1),(6,1,6,1),(7,9,1,15),(8,9,2,15),(9,9,4,17),(10,9,5,17),(11,9,6,17),(16,3,1,3),(17,4,1,4),(18,5,1,5),(19,5,1,6),(20,7,1,9),(21,10,1,10),(22,4,1,13),(23,6,1,27),(24,8,1,26),(26,3,2,3),(27,4,2,4),(28,5,2,5),(29,5,2,6),(30,6,2,27),(31,7,2,9),(32,10,2,10),(33,4,2,13),(34,8,2,26),(35,2,3,2),(36,3,3,3),(37,4,3,4),(38,5,3,5),(39,5,3,6),(40,7,3,9),(41,10,3,10),(42,4,3,13),(43,6,3,27),(44,2,4,2),(45,11,4,18),(46,12,4,19),(48,12,4,24),(49,8,4,25),(50,8,4,26),(51,6,4,27),(52,11,5,18),(53,12,5,19),(54,12,5,24),(56,8,5,25),(57,8,5,26),(58,6,5,27),(59,11,6,18),(60,12,6,19),(61,12,6,24),(63,8,6,25),(64,8,6,26),(65,6,6,27),(66,11,4,20),(67,11,5,20),(68,11,6,20),(69,15,1,32),(70,15,2,32),(71,15,3,32),(72,15,4,31),(73,15,5,31),(74,15,6,31),(75,14,1,2),(76,14,1,28),(77,14,2,2),(78,14,2,28),(79,16,4,29),(80,16,4,30),(81,16,5,29),(82,16,5,30),(83,16,6,29),(84,16,6,30),(85,17,3,28),(86,17,4,28),(87,17,5,28),(88,17,6,28),(89,2,5,2),(90,2,6,2),(91,3,4,1),(92,3,5,33),(93,3,6,33),(94,3,4,33),(95,13,4,23),(96,13,4,16),(97,13,5,23),(98,13,5,16),(99,13,6,23),(100,13,6,16);
/*!40000 ALTER TABLE `teacher_class` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `teachers`
--

DROP TABLE IF EXISTS `teachers`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
 SET character_set_client = utf8mb4 ;
CREATE TABLE `teachers` (
  `teacher_id` int(11) NOT NULL AUTO_INCREMENT,
  `name` varchar(100) DEFAULT NULL,
  `Type` varchar(45) DEFAULT NULL,
  PRIMARY KEY (`teacher_id`)
) ENGINE=InnoDB AUTO_INCREMENT=18 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci COMMENT='Holds teachers details';
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `teachers`
--

LOCK TABLES `teachers` WRITE;
/*!40000 ALTER TABLE `teachers` DISABLE KEYS */;
INSERT INTO `teachers` VALUES (1,'OBONG','Full Time'),(2,'TIMI','Full Time'),(3,'FRANCIS','Full Time'),(4,'JERRY','Full Time'),(5,'REX','Full Time'),(6,'UMOREN','Part Time'),(7,'SOLOMON','Full Time'),(8,'UFEH','Full Time'),(9,'ELDAD','Full Time'),(10,'MARVELOUS','Part Time'),(11,'ID','Full Time'),(12,'VICTOR','Full Time'),(13,'MAURICE','Full Time'),(14,'Painstil','Full Time'),(15,'Akwaima','Full Time'),(16,'Aniekan','Full Time'),(17,'Anselm','Full Time');
/*!40000 ALTER TABLE `teachers` ENABLE KEYS */;
UNLOCK TABLES;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2020-02-15 10:42:26
