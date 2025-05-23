provider "aws" {
  region = "us-east-1"
}

data "aws_iam_role" "existing_lambda_exec" {
  name = "lambda_exec_role"
}

data "aws_s3_bucket" "existing_bucket" {
  bucket = "blackjack-game-site"
}

resource "aws_s3_bucket_website_configuration" "static_site" {
  bucket = data.aws_s3_bucket.existing_bucket.id

  index_document {
    suffix = "index.html"
  }
}

resource "aws_cloudfront_distribution" "cdn" {
  origin {
    domain_name = data.aws_s3_bucket.existing_bucket.bucket_regional_domain_name
    origin_id   = "S3-blackjack-game-site"
  }
  enabled = true
  default_cache_behavior {
    allowed_methods  = ["GET", "HEAD"]
    cached_methods   = ["GET", "HEAD"]
    target_origin_id = "S3-blackjack-game-site"
    viewer_protocol_policy = "redirect-to-https"
    forwarded_values {
      query_string = false
      cookies {
        forward = "none"
      }
    }
  }
  restrictions {
    geo_restriction {
      restriction_type = "none"
    }
  }
  viewer_certificate {
    cloudfront_default_certificate = true
  }
}

resource "aws_lambda_function" "blackjack_game" {
  filename         = "blackjack_game.zip"
  function_name    = "blackjack_game"
  role             = data.aws_iam_role.existing_lambda_exec.arn
  handler          = "blackjack.handler"
  runtime          = "python3.8"
  source_code_hash = filebase64sha256("blackjack_game.zip")
}
