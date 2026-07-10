# A dummy Dockerfile with common security misconfigurations for static analysis
FROM ubuntu:20.04

# Run as root (misconfiguration)
USER root

# Run apt-get update without pinning versions or clean-up
RUN apt-get update && apt-get install -y \
    curl \
    sudo

# Expose a privileged port
EXPOSE 22

CMD ["/bin/bash"]
