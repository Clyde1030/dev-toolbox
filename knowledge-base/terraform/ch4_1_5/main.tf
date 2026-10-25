provider "google" {
  credentials = file("/Users/yu-shenglee/Desktop/Important_keys/tf101_keys.json")
  project     = "sixth-well-466802-j4"
  region      = "us-west1"
}

resource "google_storage_bucket" "quick_start_gcs" {
  name          = "quick-start-gcs-bucket-service-account"
  location      = "us-west1"
  force_destroy = true
}
