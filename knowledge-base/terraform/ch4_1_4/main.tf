provider "google" {
  project = "sixth-well-466802-j4"
  region  = "us-west1"
}

resource "google_storage_bucket" "quick_start_gcs" {
  name          = "quick-start-gcs-bucket-eason-111"
  location      = "us-west1"
  force_destroy = true
}
